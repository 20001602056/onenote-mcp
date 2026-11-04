// Shared test settings. Point the tests at a notebook/section you own:
//   ONENOTE_TEST_NOTEBOOK="My Notebook" ONENOTE_TEST_SECTION="Quick Notes" npm run test:read
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

export const NOTEBOOK = process.env.ONENOTE_TEST_NOTEBOOK;
export const SECTION = process.env.ONENOTE_TEST_SECTION;

export function requireTarget() {
  if (!NOTEBOOK || !SECTION) {
    console.error("Set ONENOTE_TEST_NOTEBOOK and ONENOTE_TEST_SECTION to a notebook and section you own.");
    process.exit(1);
  }
}

export async function connect(name = "test") {
  const c = new Client({ name, version: "1" });
  await c.connect(
    new StdioClientTransport({
      command: "node",
      args: [new URL("../server.js", import.meta.url).pathname],
      env: { ...process.env },
    })
  );
  return c;
}
