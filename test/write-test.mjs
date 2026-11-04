import { NOTEBOOK, SECTION, connect, requireTarget } from "./config.mjs";

requireTarget();
const c = await connect();
const call = async (name, args) => { const r = await c.callTool({ name, arguments: args }); console.log(`--- ${name}${r.isError ? " (ERROR)" : ""}\n${r.content[0].text.slice(0, 1500)}`); return r; };
const cr = await call("create_page", { notebook: NOTEBOOK, section: SECTION, title: "Test from onenote-mcp",
  bodyHtml: `<h2>Test page</h2><p>Created by the onenote-mcp write test on ${new Date().toISOString().slice(0, 10)}. Safe to delete.</p><pre>npm run login</pre>` });
const id = JSON.parse(cr.content[0].text).id;
await call("append_to_page", { pageId: id, html: "<p><b>Appended line</b> - append_to_page works.</p>" });
await new Promise(r => setTimeout(r, 3000));
await call("read_page", { pageId: id });
await c.close();
