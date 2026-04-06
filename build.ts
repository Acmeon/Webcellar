import url from "node:url"
import path from "node:path"

import * as rolldown from "rolldown"

// These imports enable `node --watch` to watch for changes in these modules.
import * as startup from "./bootstrap.ts"
import * as webcellar from "./webcellar.ts"

export async function build() 
{
    await rolldown.build(
    [
        {
            input: "webcellar.ts",
            output: 
            {
                file: ".webcellar/webcellar.js",
                minify: true,
                sourcemap: true
            },
        },
        {
            input: "bootstrap.ts",
            output: 
            {
                file: ".webcellar/bootstrap.js",
                minify: true,
                sourcemap: true,
            },
        },
        {
            input: url.fileURLToPath(await import.meta.resolve("es-module-shims")),
            output: 
            {
                file: ".webcellar/dependencies/es-module-shims.js",
                minify: true,
                sourcemap: true,
            },
        }
    ])
}

if(import.meta.main)  
{
    await build()
}
