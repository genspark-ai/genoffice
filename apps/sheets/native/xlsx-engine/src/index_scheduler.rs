use super::*;

pub(super) type IndexState = Arc<(Mutex<SheetIndex>, Condvar)>;
pub(super) type CellIndexJob = Box<dyn Fn(&mut dyn FnMut(usize)) + Send + Sync>;
pub(super) type MetadataJob = Box<dyn Fn() + Send + Sync>;

pub(super) struct IndexWorker {
    state: IndexState,
    index: Arc<CellIndexJob>,
    metadata: Option<MetadataJob>,
    handle: Option<JoinHandle<()>>,
    metadata_handle: Option<JoinHandle<()>>,
    started: bool,
    running: bool,
    budgeted: bool,
    dependencies: Vec<usize>,
}

impl IndexWorker {
    pub(super) fn new(
        state: IndexState,
        index: CellIndexJob,
        metadata: Option<MetadataJob>,
    ) -> Self {
        Self {
            state,
            index: Arc::new(index),
            metadata,
            handle: None,
            metadata_handle: None,
            started: false,
            running: false,
            budgeted: false,
            dependencies: Vec::new(),
        }
    }

    fn release_budget(&mut self) -> Result<(), SidecarError> {
        if self.budgeted {
            let (lock, condition) = &*self.state;
            let mut index = lock
                .lock()
                .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
            index.prefetch_bytes_remaining = None;
            self.budgeted = false;
            condition.notify_all();
        }
        Ok(())
    }
}

struct Schedule {
    workers: Vec<IndexWorker>,
    active: Option<usize>,
    viewport_ready: bool,
    budget_allocated: bool,
    reachable: Vec<bool>,
}

pub(super) struct IndexScheduler {
    schedule: Mutex<Schedule>,
    visible: Vec<usize>,
    cancelled: Arc<AtomicBool>,
    full_slots: usize,
    budget_slots: usize,
}

impl IndexScheduler {
    pub(super) fn new(
        workers: Vec<IndexWorker>,
        visible: Vec<usize>,
        cancelled: Arc<AtomicBool>,
    ) -> Arc<Self> {
        let cores = thread::available_parallelism().map_or(1, |count| count.get());
        let spare = (cores.saturating_sub(2) / 2).max(1);
        let reachable = vec![false; workers.len()];
        Arc::new(Self {
            schedule: Mutex::new(Schedule {
                workers,
                active: None,
                viewport_ready: false,
                budget_allocated: false,
                reachable,
            }),
            visible,
            cancelled,
            full_slots: spare.clamp(2, 4),
            budget_slots: spare.min(PREFETCH_CACHE_BUDGET / MIN_PREFETCH_BYTES_PER_SHEET),
        })
    }

    pub(super) fn demand(self: &Arc<Self>, sheet: usize) -> Result<(), SidecarError> {
        self.start(&mut self.schedule.lock().unwrap(), sheet, None)
    }

    pub(super) fn focus(&self, sheet: usize) {
        let mut schedule = self.schedule.lock().unwrap();
        if schedule.active != Some(sheet) {
            schedule.active = Some(sheet);
            schedule.viewport_ready = false;
        }
    }

    pub(super) fn viewport_ready(self: &Arc<Self>, sheet: usize) {
        let mut schedule = self.schedule.lock().unwrap();
        if schedule.active == Some(sheet) && !schedule.viewport_ready {
            schedule.viewport_ready = true;
            self.prefetch(&mut schedule);
        }
    }

    pub(super) fn ensure_metadata(&self, sheet: usize) {
        Self::start_metadata(&mut self.schedule.lock().unwrap().workers[sheet], sheet);
    }

    fn start_metadata(worker: &mut IndexWorker, sheet: usize) {
        if !THREADS {
            return;
        }
        let Some(job) = worker.metadata.take() else {
            return;
        };
        match thread::Builder::new()
            .name(format!("xlsx-metadata-{sheet}"))
            .spawn(job)
        {
            Ok(handle) => worker.metadata_handle = Some(handle),
            Err(error) => record_metadata_scan_failure(&worker.state, sheet, &error.into()),
        }
    }

    fn start(
        self: &Arc<Self>,
        schedule: &mut Schedule,
        sheet: usize,
        budget: Option<usize>,
    ) -> Result<(), SidecarError> {
        let worker = &mut schedule.workers[sheet];
        if worker.started {
            if budget.is_none() {
                worker.release_budget()?;
            }
            return Ok(());
        }
        worker
            .state
            .0
            .lock()
            .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?
            .prefetch_bytes_remaining = budget;
        let job = Arc::clone(&worker.index);
        if THREADS {
            let scheduler = Arc::downgrade(self);
            worker.handle = Some(
                thread::Builder::new()
                    .name(format!("xlsx-index-{sheet}"))
                    .spawn(move || {
                        job(&mut |target| {
                            if let Some(scheduler) = scheduler.upgrade() {
                                let mut schedule = scheduler.schedule.lock().unwrap();
                                schedule.workers[sheet].dependencies.push(target);
                                if schedule.reachable[sheet] {
                                    scheduler.prefetch(&mut schedule);
                                }
                            }
                        });
                        if let Some(scheduler) = scheduler.upgrade() {
                            let mut schedule = scheduler.schedule.lock().unwrap();
                            schedule.workers[sheet].running = false;
                            scheduler.prefetch(&mut schedule);
                        }
                    })?,
            );
            worker.running = true;
        } else {
            job(&mut |_| {});
        }
        worker.started = true;
        worker.budgeted = budget.is_some();
        Ok(())
    }

    fn prefetch(self: &Arc<Self>, schedule: &mut Schedule) {
        let Some(active) = schedule.active else {
            return;
        };
        if !schedule.viewport_ready || self.cancelled.load(Ordering::Acquire) {
            return;
        }

        // Walk only the active sheet's dependencies. A neighbour's formulas
        // affect new priorities if that neighbour is also a dependency.
        let mut required = vec![false; schedule.workers.len()];
        required[active] = true;
        let mut dependencies = vec![active];
        let mut cursor = 0;
        while cursor < dependencies.len() {
            for &target in &schedule.workers[dependencies[cursor]].dependencies {
                if !required[target] {
                    required[target] = true;
                    dependencies.push(target);
                }
            }
            cursor += 1;
        }
        schedule.reachable.clone_from(&required);
        let position = self
            .visible
            .iter()
            .position(|&sheet| sheet == active)
            .unwrap_or(0);
        let mut neighbours: Vec<_> = self
            .visible
            .iter()
            .copied()
            .enumerate()
            .filter(|&(_, sheet)| sheet != active)
            .collect();
        neighbours.sort_by_key(|&(index, _)| (position.abs_diff(index), index < position));
        dependencies.remove(0);
        for &(_, neighbour) in neighbours.iter().take(2) {
            if !required[neighbour] {
                required[neighbour] = true;
                dependencies.push(neighbour);
            }
        }

        let mut running = schedule
            .workers
            .iter()
            .enumerate()
            .filter(|&(sheet, worker)| sheet != active && worker.running && !worker.budgeted)
            .count();
        for sheet in dependencies {
            let worker = &schedule.workers[sheet];
            if !worker.started && running >= self.full_slots {
                return;
            }
            let additional = !worker.started || worker.running && worker.budgeted;
            if self.start(schedule, sheet, None).is_ok() {
                Self::start_metadata(&mut schedule.workers[sheet], sheet);
                running += usize::from(additional);
            }
        }

        if schedule.budget_allocated {
            return;
        }
        schedule.budget_allocated = true;
        let budgeted: Vec<_> = neighbours
            .into_iter()
            .map(|(_, sheet)| sheet)
            .filter(|&sheet| !required[sheet] && !schedule.workers[sheet].started)
            .take(self.budget_slots)
            .collect();
        let share = PREFETCH_CACHE_BUDGET.checked_div(budgeted.len());
        for sheet in budgeted {
            if self.start(schedule, sheet, share).is_ok() {
                Self::start_metadata(&mut schedule.workers[sheet], sheet);
            }
        }
    }

    pub(super) fn close(&self) {
        self.cancelled.store(true, Ordering::Release);
        let handles: Vec<_> = self
            .schedule
            .lock()
            .unwrap()
            .workers
            .iter_mut()
            .flat_map(|worker| {
                let (lock, condition) = &*worker.state;
                let _index = lock.lock();
                condition.notify_all();
                [worker.handle.take(), worker.metadata_handle.take()]
                    .into_iter()
                    .flatten()
            })
            .collect();
        for handle in handles {
            let _ = handle.join();
        }
    }

    #[cfg(test)]
    pub(super) fn with_budget(self: &Arc<Self>, sheet: usize, budget: usize) {
        self.start(&mut self.schedule.lock().unwrap(), sheet, Some(budget))
            .unwrap();
    }

    #[cfg(test)]
    pub(super) fn worker_state(&self, sheet: usize) -> (bool, bool) {
        let schedule = self.schedule.lock().unwrap();
        let worker = &schedule.workers[sheet];
        (worker.started, worker.budgeted)
    }
}
