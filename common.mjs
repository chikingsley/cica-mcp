import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
export function server() {
  return new McpServer({ name: "cica", version: "1.0.0" });
}
export function tool(mcp, name, description, inputSchema, handler) {
  mcp.registerTool(name, { description, inputSchema }, async (args) => ({
    content: [{ type: "text", text: JSON.stringify(await handler(args)) }],
  }));
}
