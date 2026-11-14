import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'

// The trusted host supplies startMcp's connection or readOnlyConnection through stdin, never a tool argument.
let input = ''
for await (const chunk of process.stdin) {
  input += chunk
  if (input.length > 64 * 1024) throw new Error('Connection configuration exceeds the size limit')
}
const connection = JSON.parse(input)
const client = new Client({ name: 'docs-http-example', version: '1' })
const transport = new StreamableHTTPClientTransport(new URL(connection.url), {
  requestInit: { headers: connection.headers },
})
try {
  await client.connect(transport)
  const tools = (await client.listTools()).tools.map((tool) => tool.name)
  const result = await client.callTool({ name: 'read_document', arguments: {} })
  if (result.isError) throw new Error('The authorized document read failed')
  console.log(JSON.stringify({ tools, result }))
} finally {
  if (transport.sessionId) await transport.terminateSession()
  await client.close()
}
