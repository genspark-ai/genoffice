/**
 * pi SDK 嵌入 smoke test
 *
 * 验证 GenOffice 工作区可以通过 file: 协议 link 到 pi 包,
 * 成功 import 并运行 createAgentSession 的最小事件流。
 *
 * 运行:
 *   cd /Users/louloulin/appx/genoffice/apps/docs && npx tsx src/renderer/ai/pi-smoke.ts
 *
 * 看到 "Events: [message_start, message_update, ..., agent_end]" 即成功。
 */

import { createAgentSession, ModelRuntime, SessionManager } from "@earendil-works/pi-coding-agent";

async function main() {
  console.log("[pi-smoke] start, importing pi SDK…");
  const modelRuntime = await ModelRuntime.create();
  console.log("[pi-smoke] modelRuntime created");

  const { session } = await createAgentSession({
    sessionManager: SessionManager.inMemory(),
    modelRuntime,
  });
  console.log("[pi-smoke] session created");

  const events: string[] = [];
  session.subscribe((event) => {
    events.push(event.type);
  });

  console.log("[pi-smoke] sending prompt…");
  await session.prompt("Say hi in 5 words.");
  console.log("[pi-smoke] prompt finished");
  console.log("Events:", events);

  session.dispose();
  console.log("[pi-smoke] OK — pi SDK embedding works from GenOffice");
}

main().catch((err) => {
  console.error("[pi-smoke] FAILED:", err);
  process.exit(1);
});
