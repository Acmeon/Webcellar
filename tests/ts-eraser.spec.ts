// Start Webcellar server first. Serve files from ./demo

import path from "node:path"
import url from "node:url"
import fs from "fs/promises"
import { test, expect } from '@playwright/test';

var cwd = path.dirname(path.dirname(url.fileURLToPath(import.meta.url)).replaceAll("\\", "/"))

// TS START
var source = () => 
{
    // TODO: Systematically and comprehensively test TypeScript annotations.
    var v0: number = 0
    var v1 = 0 as number
    var {v2, v3}: {v2: number, v3: string} = {v2: 10, v3: ""}
    var {v4, v5}: {[key: string]: number} = {v4: 10, v5: 11}
    var [v6, v7]: number[] = [1, 2]
    var v8 = new Map<string, string>()
    var v9 = new Map<number, Set<number>>()

    function f0(a0: number, a1: number | string, a2: boolean[]){}

    class c0 
    {
        a: string = ""
        b: number

        constructor(v: number = 10)
        {
            this.b = v
        }
    }
}
// TS END

var ts = (await fs.readFile(import.meta.filename, "utf-8")).split("// TS START")[1].split("// TS END")[0] 


test("ts1", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)

    var result = await page.evaluate(async (ts) => 
    {
        var bs = await import("/.webcellar/bootstrap.js")
        var js = bs.transform(ts, "")

        try
        {
            eval(js)
            return {js, ok: true}
        }
        catch
        {
            return {js, ok: false}
        }
    }, ts)

    console.log(result.js)

    expect(result.ok).toBe(true)
})