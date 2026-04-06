// Start Webcellar server first. Serve files from ./demo

import path from "node:path"
import url from "node:url"
import { test, expect } from '@playwright/test';

var cwd = path.dirname(path.dirname(url.fileURLToPath(import.meta.url)).replaceAll("\\", "/"))

test("req1", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)
    await page.evaluate((cwd) => {(window as any).cwd = cwd}, cwd)

    var result = await page.evaluate(async () => 
    {
        var res0 = await fetch(`/${cwd}/demo/demo.xlsx.ts`)
        
        return res0.ok
    })

    expect(result).toBe(true)
})

test("req2", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)
    await page.evaluate((cwd) => {(window as any).cwd = cwd}, cwd)

    var result = await page.evaluate(async () => 
    {
        var res0 = await fetch(`/${cwd}/demo/module.ts`)
        
        return !res0.ok
    })

    expect(result).toBe(true)
})

test("req3", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)
    await page.evaluate((cwd) => {(window as any).cwd = cwd}, cwd)

    var result = await page.evaluate(async () => 
    {
        var res0 = await fetch(`/${cwd}/demo/foo/mod.js`)
        
        return !res0.ok
    })

    expect(result).toBe(true)
})

test("req4", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)
    await page.evaluate((cwd) => {(window as any).cwd = cwd}, cwd)

    var result = await page.evaluate(async () => 
    {
        var res0 = await fetch(`/${cwd}/demo/bar/data.json`)
        
        return !res0.ok
    })

    expect(result).toBe(true)
})

test("req5", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)
    await page.evaluate((cwd) => {(window as any).cwd = cwd}, cwd)

    var result = await page.evaluate(async () => 
    {
        var res0 = await fetch(`/${cwd}/demo/baz/data.json`)
        
        return !res0.ok
    })

    expect(result).toBe(true)
})

test("req6", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)
    await page.evaluate((cwd) => {(window as any).cwd = cwd}, cwd)

    var result = await page.evaluate(async () => 
    {
        var res0 = await fetch(`/${cwd}/demo/demo.xlsx.ts`)
        var res1 = await fetch(`/${cwd}/demo/module.ts`)
        
        return res0.ok && res1.ok
    })

    expect(result).toBe(true)
})

test("req7", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)
    await page.evaluate((cwd) => {(window as any).cwd = cwd}, cwd)

    var result = await page.evaluate(async () => 
    {
        var res0 = await fetch(`/${cwd}/demo/module.ts`)
        var res1 = await fetch(`/${cwd}/demo/demo.xlsx.ts`)
        
        return !res0.ok && res1.ok
    })

    expect(result).toBe(true)
})

test("req8", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)
    await page.evaluate((cwd) => {(window as any).cwd = cwd}, cwd)

    var result = await page.evaluate(async () => 
    {
        var res0 = await fetch(`/${cwd}/demo/demo.xlsx.ts`)
        var res1 = await fetch(`/${cwd}/demo/foo/mod.js`)
        
        return res0.ok && res1.ok
    })

    expect(result).toBe(true)
})

test("req9", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)
    await page.evaluate((cwd) => {(window as any).cwd = cwd}, cwd)

    var result = await page.evaluate(async () => 
    {
        var res0 = await fetch(`/${cwd}/demo/demo.xlsx.ts`)
        var res1 = await fetch(`/${cwd}/demo/bar/data.json`)
        
        return res0.ok && !res1.ok
    })

    expect(result).toBe(true)
})

test("req10", async ({page}) => 
{
    await page.goto(`https://localhost:29640/.webcellar/help.html`)
    await page.evaluate((cwd) => {(window as any).cwd = cwd}, cwd)

    var result = await page.evaluate(async () => 
    {
        var res0 = await fetch(`/${cwd}/demo/demo.xlsx.ts`)
        var res1 = await fetch(`/${cwd}/demo/baz/data.json`)
        
        return res0.ok && res1.ok
    })

    expect(result).toBe(true)
})