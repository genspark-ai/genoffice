# GenOffice

An offline-first office suite: the original file is the source of truth, edits are
narrow patches, and everything the editor did not touch survives the round trip.
The language below covers the AI panel, which is the part of the product where a
document and a conversation meet.

## Language

### The AI panel

**Agent panel**:
The conversation surface beside a document, where the user asks the model to read or
change the open file.
_Avoid_: Chat sidebar, copilot, assistant pane

**Composer**:
The input box at the bottom of the agent panel: one draft, one send.
_Avoid_: Input bar, text field, prompt box

**Draft**:
The text currently in the composer, before it becomes a message.
_Avoid_: Prompt, input, query

**Message**:
A draft that has been sent and now belongs to the conversation.
_Avoid_: Turn, query, request

**Scope**:
The part of the document a message is allowed to act on, attached to the draft.
_Avoid_: Selection, context, target

### Voice input

**Voice input**:
Speaking a draft instead of typing it. The speech becomes draft text; it never becomes
a message on its own.
_Avoid_: Dictation, speech-to-text, STT, voice typing

**Voice session**:
One span of listening, from the microphone opening to the transcript being committed
or thrown away.
_Avoid_: Recording, take, utterance

**Transcript**:
The text a voice session produces.
_Avoid_: Result, output, recognition

**Interim transcript**:
The revisable transcript shown while the user is still speaking. It is expected to be
replaced, and it is never sent.
_Avoid_: Partial, streaming result, live text

**Final transcript**:
The transcript a voice session settles on when it ends. This is what the draft keeps.
_Avoid_: Committed text, result

**Transcriber**:
The component that turns captured audio into a transcript.
_Avoid_: Engine, model, ASR, recognizer

**Anchored session**:
A voice session whose interim transcript occupies a known span of the draft.
_Avoid_: Attached, bound

**Detached session**:
A voice session that has lost its span of the draft because the user typed during it;
the session ends, the interim transcript already in the draft is kept, and the final
transcript is dropped rather than written over what the user typed.
_Avoid_: Broken, fallback

**Discarded session**:
A voice session the user cancelled. Its transcript never reaches the draft.
_Avoid_: Cancelled, aborted
