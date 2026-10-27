# Audio never leaves the device

The suite's promise is that only AI calls leave the machine, and an audio recording is
a strictly more sensitive artefact than the text it turns into. Voice input therefore
runs on a local transcriber: the microphone stream is captured in the renderer, handed
to an on-device engine, and dropped. No audio is uploaded, persisted to disk, or
attached to a message.

The transcriber is an interface (`VoiceTranscriber`) rather than a call site, so a
cloud backend can be added later behind an explicit setting — but it must be opt-in,
and the default must stay local. Because the audio never leaves, the operating
system's microphone prompt is the only consent step; there is no in-app consent
dialog to dismiss.

## Consequences

- Transcription cost is paid in CPU and disk (a model file), not in tokens.
- The engine is delivered separately from this UI work: the composer slot, the
  session state machine and the draft splicing land first and are testable without
  any model present.
- Any future cloud transcriber has to be visible in settings and off by default, or
  this decision has been reversed silently.
