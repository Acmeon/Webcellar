#!/usr/bin/env node

import express from "express"
import {fileURLToPath} from "node:url"
import path from "node:path"
import fs from "fs/promises"
import crypto from "crypto"
import cookieParser from "cookie-parser"
import * as esModuleLexer from "es-module-lexer"
import {Command} from "commander"
import https from "https"
import officeAddinDevCerts from "office-addin-dev-certs"
import officeAddinDebugging, {AppType} from "office-addin-debugging"

import * as build from "./build.ts"

async function isPathAccessAuthorized(fp: string)
{
    try
    {
        await fs.access(fp)
        
        if((await fs.stat(fp)).isDirectory())
        {
            return false   
        }

        for(var glob of globs)
        {
            if(path.matchesGlob(fp, glob))
            {
                return true   
            }   
        }
    }
    catch{}

    return false
}

var cwd = path.dirname(fileURLToPath(import.meta.url)).replaceAll("\\", "/")

console.log(
`
Webcellar 
=========

Installation directory:
${cwd}
`)

var program = new Command()
program.name("webcellar")
       .argument("<dirs...>", "directories (one or more) from which files are served (for example, use C:/ on Windows to grant access to all files under the C drive)")
       .option("--mode <mode>", "execution mode for Webcellar: 'init' initializes, 'run' starts the server, 'deinit' removes it; by default '', which executes 'init' (if needed) and then 'run'")
       .option("--content-security-policy-sources <sources...>", "additional sources for the Content Security Policy (CSP) default-src directive (allowed by default: 'self', blob:, data:, 'unsafe-inline', 'unsafe-eval')")
       .exitOverride((e) => 
        {
            if(e.code != "commander.helpDisplayed")
            {
                console.log("Invalid command, see the README.md file for reference (https://github.com/Acmeon/Webcellar)")
                console.log("")
                console.log("For example, to serve files under the C drive on Windows, run the following command:")
                console.log(`npx webcellar C:/`)
                console.log("")
            }
        })
        .parse(process.argv)


var globs = [...program.args.map(p => path.resolve(`${p}/**`).replaceAll("\\", "/")), path.resolve(`${cwd}/.webcellar/**`).replaceAll("\\", "/")]
var csps = ["blob:", "data:", "'self'", "'unsafe-inline'", "'unsafe-eval'", ...(program.opts().contentSecurityPolicySources as string[] ?? [])] 
var mode = program.opts().mode ?? ""
var modeOption = program.opts().mode != null

var webcellarCookieToDir = new Map<string, string>()

if(mode == "")
{
    // Contains a README.md file by default.
    if((await fs.readdir(`${cwd}/.webcellar/dependencies/`)).length <= 1)
    {
        mode = "init" 
    }
    else
    {
        mode = "run"
    }
}


// Initialize mode
if(mode == "init")
{
    await officeAddinDebugging.startDebugging(`${cwd}/.webcellar/manifest.xml`, 
    {
        appType: AppType.Desktop, 
        enableLiveReload: false,
        enableSideload: true,
        enableDebugging: false
    })

    await build.build()
}
else if(mode == "run")
{
    // Nothing to init
}
else if(mode == "deinit")
{
    await officeAddinDebugging.stopDebugging(`${cwd}/.webcellar/manifest.xml`)
    process.exit()   
}


// Run server
var app = express()
var options = await officeAddinDevCerts.getHttpsServerOptions(365)
var server = https.createServer(options, app).listen(29640)

app.use(cookieParser())
app.use(express.json())


console.log(
`
Webcellar is serving files on: 
https://localhost:29640

Serving files that match at least one of the glob patterns: 
${globs.join("\n")}

Content Security Policy sources:
${csps.join("\n")}
`)

app.use(async (req, res) => 
{
    console.log()
    console.log(`URL: ${decodeURIComponent(req.url)}`)

    if(mode == "init")
    {
        // Cache cannot be zero, otherwise Excel fails to display the Webcellar icon.
        res.setHeader("Cache-Control", "public, max-age=1, must-revalidate")
        
        if(req.path == "/")
        {
            if(req.body)
            {
                var officeUrls = (req.body as string[]).filter(url => !url.startsWith("https://localhost:29640"))
    
                // Download office.js files.
                // The main file is hosted at https://appsforoffice.microsoft.com/lib/1/hosted/office.js
                var error = false
                for(var url of officeUrls)
                {
                    try
                    {   
                        console.log(`Downloading: ${url}`)
    
                        var code = await (await fetch(url)).text()
    
                        console.log(`Download successful.`)
    
                        var fp = `${cwd}/.webcellar/dependencies/office-js/${url.replace("https://appsforoffice.microsoft.com/lib/1/hosted/", "")}`
    
                        console.log(`Writing file to: ${fp}`)
    
                        await fs.mkdir(path.dirname(fp), {recursive: true})
                        await fs.writeFile(fp, code)
    
                        console.log(`Writing successful.`)
                    }
                    catch
                    {
                        error = true
                        console.log(`ERROR!`)
                    }
                }
    
                if(!error)
                {
                    console.log("")
                    console.log("Initialization of Webcellar completed successfully.")
                    console.log("Please close the running Excel instance.")

                    res.send("OK")

                    if(modeOption)
                    {
                        process.exit()   
                    }
                    else
                    {
                        mode = "run"
                    }
                }
                else
                {
                    console.log("")
                    console.log("Initialization of Webcellar failed.")
                    console.log("Please close the running Excel instance.")

                    res.send("Failed")

                    process.exit()

                }
            }
        }
        else if(req.path.startsWith("/.webcellar"))
        {
            // Serve Webcellar files.
            
            var req_path = path.join(cwd, req.path)
    
            // Serve special taskpane for initialization.
            if(req.path == "/.webcellar/taskpane.html")
            {
                req_path = path.join(cwd, "/.webcellar/taskpane-init.html")
            }
    
            if(!(await isPathAccessAuthorized(req_path)))
            {
                console.log(`Response: Internal Webcellar file not found.`)
    
                res.status(404).send("Not found.")
            }
    
            console.log("Response: Serve internal Webcellar file.")
    
            res.sendFile(req_path, {dotfiles: "allow"})
        }
        else
        {
            res.status(404).send("Not found.")
        }
    }
    else if(mode == "run")
    {
        res.setHeader("Content-Security-Policy", `default-src ${csps.join(" ")}`)
        res.setHeader('Cache-Control', "no-store, max-age=0, must-revalidate")

        if(req.path.startsWith("/.webcellar"))
        {
            // Serve Webcellar files.

            var reqFilePath = path.join(cwd, req.path)

            if(!(await isPathAccessAuthorized(reqFilePath)))
            {
                console.log(`Response: Internal Webcellar file not found.`)

                res.status(404).send("Not found.")
                return
            }

            console.log("Response: Serve internal Webcellar file.")
            res.sendFile(reqFilePath, {dotfiles: "allow"})
        }
        else 
        {
            // Serve from the file system, if authorized.

            // Remove leading / from path on Windows (it is always present in URL after domain)
            var reqFilePath = process.platform == "win32" ? path.resolve(decodeURIComponent(req.path).slice(1)) : path.resolve(decodeURIComponent(req.path))

            if(!(await isPathAccessAuthorized(reqFilePath)))
            {
                console.log(`Response: Not found.`)

                res.status(404).send("Not found.")
                return
            }

            if(reqFilePath.endsWith(".xlsx.js") || reqFilePath.endsWith(".xlsx.ts"))
            {
                // Direct access to a Webcellar file.

                var webcellarFilePaths = [reqFilePath]

                // Analyze static imports from the requested Webcellar file.
                // This is required because file access authorization is handled with cookies,
                // which must be set before request. However, module fetch order should 
                // not be relied upon, thus, the cookies corresponding to directly imported 
                // Webcellar files will be set immediately.
                var contents = await fs.readFile(reqFilePath, {encoding: "utf-8"})
                try
                {
                    var imports = esModuleLexer.parse(contents)[0]
                
                    for(var imp of imports)
                    {   
                        if(imp.n?.endsWith(".xlsx.js") || imp.n?.endsWith(".xlsx.ts"))
                        {
                            var webcellarFilePath = path.resolve(path.dirname(reqFilePath), imp.n)

                            if(await isPathAccessAuthorized(webcellarFilePath))
                            {
                                webcellarFilePaths.push(webcellarFilePath)
                            }
                        }
                    }
                }
                catch{}

                webcellarFilePaths = Array.from(new Set(webcellarFilePaths))

                var redirect = false

                for(var webcellarFilePath of webcellarFilePaths)
                {
                    var webcellarDirPath = path.dirname(webcellarFilePath)

                    // Each Webcellar dir is associated with a unique cookie
                    var webcellarDirCookieName = `Webcellar#${encodeURIComponent(webcellarDirPath)}`
                    var webcellarDirCookieVal = req.cookies[webcellarDirCookieName] as string ?? ""

                    if(webcellarDirCookieVal == "" || webcellarCookieToDir.get(webcellarDirCookieVal) != webcellarDirPath)
                    {
                        webcellarDirCookieVal = crypto.randomBytes(32).toString("hex")

                        webcellarCookieToDir.set(webcellarDirCookieVal, webcellarDirPath)

                        res.cookie(webcellarDirCookieName, webcellarDirCookieVal,
                        {
                            path: webcellarDirPath.replace("\\", "/"),
                            httpOnly: true,
                            secure: true,
                            sameSite: "strict"
                        })

                        redirect = true
                    }
                }

                if(redirect)
                {
                    console.log(`Response: Set cookies and redirect to Webcellar root file.`)
                    res.redirect(302, req.path)
                }
                else
                {
                    console.log(`Response: Serve file.`)
                    res.sendFile(reqFilePath, {dotfiles: "allow"})
                }
            }
            else
            {
                // If authorized, serve the request file.

                // Determine nearest Webcellar root dir (note that the nearest dir does not depend on served directories). 
                var webcellarDirPath = path.dirname(reqFilePath)
                while(true)
                {
                    if((await Array.fromAsync(fs.glob(`${webcellarDirPath}/*.xlsx.{js,ts}`))).length > 0)
                    {
                        // Webcellar root dir found
                        break
                    }

                    var next = path.dirname(webcellarDirPath)

                    if(next == webcellarDirPath)
                    {
                        webcellarDirPath = ""
                        break   
                    }
                    else
                    {
                        webcellarDirPath = next
                    }
                }

                // Each Webcellar root dir is associated with a unique cookie
                var webcellarDirCookieName = `Webcellar#${encodeURIComponent(webcellarDirPath)}`
                var webcellarDirCookieVal = req.cookies[webcellarDirCookieName] as string ?? ""

                if(webcellarDirPath == "" || webcellarDirCookieVal == "" || webcellarCookieToDir.get(webcellarDirCookieVal) != webcellarDirPath)
                {
                    // Not authorized
                    console.log("Response: Not authorized.")
                    res.status(404).send("Not found.")
                }
                else
                {
                    // Authorized
                    console.log("Response: Authorized, serve file.")
                    res.sendFile(reqFilePath, {dotfiles: "allow"})
                }
            }
        }
    }
    else if(mode == "deinit")
    {
        // Should not happen
        process.exit()   
    }
})