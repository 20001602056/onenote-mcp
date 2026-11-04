import { NOTEBOOK, SECTION, connect, requireTarget } from "./config.mjs";

requireTarget();
const c = await connect();
console.log((await c.listTools()).tools.map(t => t.name).join(", "));
const nb = await c.callTool({ name: "list_notebooks", arguments: {} });
console.log(nb.content[0].text.slice(0, 1500));
const pg = await c.callTool({ name: "list_pages", arguments: { notebook: NOTEBOOK, section: SECTION } });
console.log(pg.isError ? "ERR " : "", pg.content[0].text.slice(0, 2500));
await c.close();
