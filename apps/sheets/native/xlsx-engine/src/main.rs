use std::io::{self, BufRead, BufReader, BufWriter, Write};
use std::sync::{Arc, Mutex, mpsc};
use std::thread;

use xlsx_sidecar::THREADS;
use xlsx_sidecar::WorkbookSessions;
use xlsx_sidecar::protocol::{
    CancelledRequests, Command, Engine, InFlightRequest, PROTOCOL_VERSION, RecalcWorker, Request,
    Response, SharedOutput, handle_request, write_response,
};

fn main() {
    if !THREADS {
        serve_sequentially();
        return;
    }
    let stdin = io::stdin();
    let output: SharedOutput = Arc::new(Mutex::new(BufWriter::new(io::stdout())));
    let mut sessions = WorkbookSessions::new();
    let recalc = RecalcWorker::new();
    let cancelled = CancelledRequests::new();
    let in_flight = InFlightRequest::new();

    // Requests execute on this thread in arrival order; a dedicated reader
    // drains stdin so a cancel can jump the queue while earlier requests are
    // still pending. The small bound keeps stdin backpressure close to the
    // old read-one-handle-one loop.
    let (queue, requests) = mpsc::sync_channel::<Request>(8);
    let reader_output = Arc::clone(&output);
    let reader_cancelled = cancelled.clone();
    let reader_in_flight = in_flight.clone();
    let reader = thread::Builder::new()
        .name("xlsx-reader".into())
        .spawn(move || {
            for line in BufReader::new(stdin.lock()).lines() {
                let line = match line {
                    Ok(line) => line,
                    Err(error) => {
                        let _ = write_response(
                            &reader_output,
                            &Response::failure(
                                String::new(),
                                "stdin_error",
                                format!("Unable to read sidecar request: {error}"),
                            ),
                        );
                        break;
                    }
                };
                let request: Request = match serde_json::from_str(&line) {
                    Ok(request) => request,
                    Err(error) => {
                        if write_response(
                            &reader_output,
                            &Response::failure(
                                String::new(),
                                "invalid_json",
                                format!("Invalid sidecar request: {error}"),
                            ),
                        )
                        .is_err()
                        {
                            break;
                        }
                        continue;
                    }
                };
                if let Command::Cancel { target_request_id } = request.command {
                    let response = if request.version == PROTOCOL_VERSION {
                        reader_cancelled.insert(target_request_id.clone());
                        // A target already running is past the queue's reach;
                        // trip the flag its open is polling.
                        reader_in_flight.cancel(&target_request_id);
                        Response::success(
                            request.request_id,
                            serde_json::json!({ "cancelled": true }),
                        )
                    } else {
                        Response::failure(
                            request.request_id,
                            "unsupported_version",
                            "Unsupported sidecar protocol version.".into(),
                        )
                    };
                    if write_response(&reader_output, &response).is_err() {
                        break;
                    }
                    continue;
                }
                if queue.send(request).is_err() {
                    break;
                }
            }
        });
    if reader.is_err() {
        return;
    }

    for request in requests {
        in_flight.begin(&request.request_id);
        let response = handle_request(
            request,
            &mut sessions,
            &recalc,
            &cancelled,
            &in_flight,
            &output,
        );
        if let Some(response) = response
            && write_response(&output, &response).is_err()
        {
            break;
        }
    }
}

/// Without threads there is no reader to let a cancel jump the queue: each
/// request runs to completion before the next line is read.
fn serve_sequentially() {
    let mut engine = Engine::new();
    let mut output = BufWriter::new(io::stdout());
    for line in io::stdin().lock().lines() {
        let Ok(line) = line else { break };
        if let Some(response) = engine.handle_line(&line)
            && (writeln!(output, "{response}").is_err() || output.flush().is_err())
        {
            break;
        }
    }
}
