#!/usr/bin/env node

import express from "express"
import {fileURLToPath} from "node:url"
import path from "node:path"
import fss from "fs"
import fs from "fs/promises"
import crypto from "crypto"
import cookieParser from "cookie-parser"
import * as esModuleLexer from "es-module-lexer"
import {Command} from "commander"
import https from "https"
import officeAddinDevCerts from "office-addin-dev-certs"
import officeAddinDebugging, {AppType} from "office-addin-debugging"
import * as rolldown from "rolldown"

// These imports enable `node --watch` to watch for changes in these modules.
import * as startup from "./bootstrap.ts"
import * as webcellar from "./webcellar.ts"

var cwd = path.dirname(fileURLToPath(import.meta.url)).replaceAll("\\", "/")

export async function build() 
{
    await rolldown.build(
    [
        {
            // Build start.ts, bc "Stripping types is currently unsupported for files under node_modules"...
            input: `${cwd}/start.ts`,
            output: 
            {
                file: `${cwd}/index.js`,
                minify: false,
                sourcemap: false,
                codeSplitting: false,
            },
            external: (id) => 
            {
                return !id.endsWith(".ts")
            }
        },
        {
            input: `${cwd}/webcellar.ts`,
            output: 
            {
                file: `${cwd}/.webcellar/webcellar.js`,
                minify: true,
                sourcemap: true
            },
        },
        {
            input: `${cwd}/bootstrap.ts`,
            output: 
            {
                file: `${cwd}/.webcellar/bootstrap.js`,
                minify: true,
                sourcemap: true,
            },
        },
        {
            input: fileURLToPath(await import.meta.resolve("es-module-shims")),
            output: 
            {
                file: `${cwd}/.webcellar/dependencies/es-module-shims.js`,
                minify: true,
                sourcemap: true,
            },
        }
    ])
}

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


console.log(
`
Webcellar 
=========

Installation directory:
${cwd}
`)

var program = new Command()
program.name("webcellar")
       .argument("[dirs...]", "directories from which files are served (for example, use C:/ on Windows to grant access to all files under the C drive)")
       .option("--mode <modes...>", "execution mode for Webcellar: 'init' initializes, 'build' builds dependencies, 'run' starts the server, 'deinit' removes it; by default 'init' (if needed), 'build' (if needed) and then 'run'")
       .option("--content-security-policy-sources <sources...>", "additional sources for the Content Security Policy (CSP) default-src directive (allowed by default: 'self', blob:, data:, 'unsafe-inline', 'unsafe-eval')")
       .exitOverride((e) => 
        {
            if(e.code != "commander.helpDisplayed")
            {
                console.log("Invalid command, see the README.md file for reference (https://github.com/Acmeon/Webcellar)")
            }
        })
        .parse(process.argv)


var globs = [...program.args.map(p => path.resolve(`${p}/**`).replaceAll("\\", "/")), path.resolve(`${cwd}/.webcellar/**`).replaceAll("\\", "/")]
var csps = ["blob:", "data:", "'self'", "'unsafe-inline'", "'unsafe-eval'", ...(program.opts().contentSecurityPolicySources as string[] ?? [])] 
var modes = program.opts().mode as string[] ?? []

if(modes.length == 0)
{
    if(!fss.existsSync(`${cwd}/.webcellar/dependencies/office-js/init.txt`))
    {
        modes.push("init") 
    }

    if(!fss.existsSync(`${cwd}/.webcellar/dependencies/build.txt`))
    {
        modes.push("build") 
    }

    modes.push("run")
}


console.log(`Modes:`)
console.log(modes.join(", "))
console.log("")

if(modes.includes("run") && globs.length <= 1)
{
    console.log("Error!")
    console.log("Current execution modes include 'run', but no directories have been specified.")
    console.log("")
    console.log("For example, to serve files under the C drive on Windows, run the following command:")
    console.log(`npx webcellar C:/`)
    process.exit()
}


// Initialize
if(modes.includes("init"))
{
    await officeAddinDebugging.startDebugging(`${cwd}/.webcellar/manifest.xml`, 
    {
        appType: AppType.Desktop, 
        enableLiveReload: false,
        enableSideload: true,
        enableDebugging: false
    })

    await fs.mkdir(`${cwd}/.webcellar/dependencies/office-js/`, {recursive: true})
    await fs.writeFile(`${cwd}/.webcellar/dependencies/office-js/init.txt`, "init", "utf8")
}

if(modes.includes("build"))
{
    await build()
    
    await fs.mkdir(`${cwd}/.webcellar/dependencies/`, {recursive: true})
    await fs.writeFile(`${cwd}/.webcellar/dependencies/build.txt`, "build", "utf8")
}

if(modes.includes("run"))
{
    // Nothing to init
}

if(modes.includes("deinit"))
{
    await officeAddinDebugging.stopDebugging(`${cwd}/.webcellar/manifest.xml`)
}



// Start server

if(!(modes.includes("init") || modes.includes("run")))
{
    process.exit() 
}

var app = express()
var options = await officeAddinDevCerts.getHttpsServerOptions(365)
var server = https.createServer(options, app).listen(29640)
var webcellarCookieToDir = new Map<string, string>()

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

    if(modes.includes("init"))
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
                }
                else
                {
                    console.log("")
                    console.log("Initialization of Webcellar failed.")
                    console.log("Please close the running Excel instance.")

                    res.send("Failed")
                }

                modes = modes.filter(m => m != "init")
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
    else if(modes.includes("run"))
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
    else
    {
        // Should not happen
        process.exit()   
    }
})
