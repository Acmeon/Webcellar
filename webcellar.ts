import * as utils from "./utils.ts"

// TODO: Implement (consider input and output)
// class RawValue
// {
//     value

//     constructor(value: any)
//     {
//         this.value = value
//     }
// }

// export function raw(value: any)
// {
//     return new RawValue(value)
// }

class Meta 
{
    input?: "convert" | "raw" = "convert"
    output?: "convert" | "raw" = "convert"
    
    // TODO: Handle all properties in metadata reference (https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-json#metadata-reference).
    excel?: 
    {
        [key: string]: any

        id?: string
        name?: string
        options?: any
        parameters?: 
        {
            [key: string]: any

            name: string
            description?: string
            dimensionality?: "scalar" | "matrix"
            type?: "boolean" | "number" | "string" | "any"
        }[]
        result?: 
        {
            dimensionality: "scalar" | "matrix"
            type?: "boolean" | "number" | "string" | "any"
        }

    } = {}
}


var metadata = new Map<any, Meta>()
export function meta(value: any, meta: Meta): void
{
    metadata.set(value, meta)
}

export function input(value: any, root = false): any
{
    // Convert Excel input (e.g., entities etc) to JS values
    /*
        Conversion rules:
        - Excel string => JS string
        - Excel number => JS number
        - Excel boolean => JS boolean
        - Excel entity => JS object
        - Excel 1x1 matrix<any> => JS any
        - Excel 1xn (row) matrix<any> => JS array<any> (1 dimensional)
        - Excel nx1 (col) matrix<any> => JS array<any> (1 dimensional)
        - Excel nxm matrix<any> => JS array<array<any>>
    */

    if(root)
    {
        if(Array.isArray(value))
        {
            if(value.length === 1 && Array.isArray(value[0]) && value[0].length === 1) 
            {
                // 1x1 matrix [[v00]]
                return input(value[0][0])
            }
            else if(value.length === 1 && Array.isArray(value[0]) && value[0].length > 1) 
            {
                // 1xn matrix [[v00, ..., v0m]]

                var arr = []
                for(let v of value[0])
                {
                    arr.push(input(v))
                }
                return arr
            }
            else if(value.length > 1 && value.every(row => Array.isArray(row) && row.length === 1)) 
            {
                // nx1 matrix: [[v00], [v10], ..., [vn0]]

                var arr = []
                for(let v of value)
                {
                    arr.push(input(v[0]))
                }
                return arr
            }
            else if(value.length > 1 && value.every(r => Array.isArray(r)) && value.every(r => r.length === value[0].length) && value[0].length > 1)
            {
                // nxm matrix: [[v00, ..., v0m], [v10, ..., v1m], ..., [vn0, ..., vnm]]

                var arr = []
                for(let v of value)
                {
                    var nested = []
                    for(let vv of v)
                    {
                        nested.push(input(vv))
                    }

                    arr.push(nested)
                }

                return arr
            }
            else if(value.every(v => !Array.isArray(v)))
            {
                // plain array: [v0, v1, ..., vn]
                
                var arr = []

                for(let v of value)
                {
                    arr.push(input(v))   
                }

                return arr
            }
            else
            {
                // fallback, e.g., "matrix" where dimensions do not match
                return null
            }
        }
        else
        {
            return input(value)
        }
    }
    else
    {
        // Most of these are probably unnecessary.

        if(value == null)
        {
            return null
        }
        else if(typeof(value) == "string")
        {
            return value
        }
        else if(typeof(value) == "number")
        {
            return value
        }
        else if(typeof(value) == "bigint")
        {
            return value
        }
        else if(typeof(value) == "boolean")
        {
            return value
        }
        else if(typeof(value) == "symbol")
        {
            return value 
        }
        else if(typeof(value) == "undefined")
        {
            return value 
        }
        else if(typeof(value) == "object")
        {
            var cv: Excel.CellValue = value

            if(cv.type == "String" || cv.type == "Double" || cv.type == "Boolean")
            {
                return cv.basicValue   
            }
            else if(cv.type == "Entity")
            {
                // TODO: Handle other cases (e.g., Map, Set).

                if(cv.text == "Array")
                {
                    var arr: any[] = []

                    if(cv.properties?.data)
                    {
                        for(var i = 0; i < value.properties.data.elements[0].length; i++)
                        {
                            arr.push(input(value.properties.data.elements[0][i]))
                        }
                    }
                    
                    return arr
                }
                else
                {
                    // TODO: Set prototype.

                    var obj: any = {}

                    for(var key in cv.properties)
                    {
                        obj[key] = input(cv.properties[key])
                    }

                    return obj
                }
            }
            else
            {
                return null
            }
        }
        else if(typeof(value) == "function")
        {
            return value
        }
        else
        {
            return null
        }
    }

}

export function output(value: any, root = false): any
{
    // TODO: Implement other cases.
    // TODO: implement raw (which would allow, e.g., formatted numbers)? Use properties.provider.description as raw indicator?

    // handle here, bc if handled in return (by checking if root is array entity), then how to differentiate between matrix and matrix as obj?
    // scalar => scalar
    // matrix => matrix
    // object => object
    // matrix as object => object (webcellar.converter.encode([1, 2, 3]))
    // matrix as object => object (webcellar.output([1, 2, 3]))

    // webcellar.output([1, 2, 3], true) => [[1, 2, 3]]
    // webcellar.output([1, 2, 3]) => {type: "Entity", ...}

    // if(value instanceof RawValue)
    // {
    //     return value.value   
    // }

    if(root)
    {
        if(Array.isArray(value))
        {
            var arr = []
            var matrix = false

            // TODO: Check if all entries are arrays of same length, which would imply a matrix?
            for(var v of value)
            {
                if(Array.isArray(v))
                {
                    var nested = []
                    for(var vv of v)
                    {
                        nested.push(output(vv))
                    }

                    arr.push(nested)

                    matrix = true
                }
                else
                {
                    arr.push(output(v))
                }
            }

            if(matrix)
            {
                return arr   
            }
            else
            {
                return [arr]
            }
        }   
        else
        {
            return [[output(value)]]
        }
    }
    else
    {
        if(value == null)
        {
            return null
        }
        else if(typeof(value) == "string")
        {
            return {type: "String", basicValue: value}   
        }
        else if(typeof(value) == "number")
        {
            return {type: "Double", basicValue: value}
        }
        else if(typeof(value) == "bigint")
        {
            return null
        }
        else if(typeof(value) == "boolean")
        {
            return {type: "Boolean", basicValue: value}   
        }
        else if(typeof(value) == "symbol")
        {
            return null 
        }
        else if(typeof(value) == "undefined")
        {
            return null 
        }
        else if(typeof(value) == "object")
        {
            var ctor = value.constructor

            // TODO: Implement other "special" cases, e.g., Map

            if(ctor == Array)
            {
                // TODO: Set text to similar as printing in Chrome devtools, e.g., [1, 2, 3, ...]
                
                let elements: any[][] = [[]]

                let entity: Excel.EntityCellValue = 
                {
                    type: "Entity",
                    text: "Array",
                    properties: 
                    {
                        data: 
                        {
                            type: "Array",
                            elements: elements
                        },
                        length: value.length,
                    }
                }

                for(var v of value)
                {
                    elements[0].push(output(v))
                }

                return entity
            }
            else
            {
                // TODO: Set text to similar as printing in Chrome devtools, e.g., {...}?
                // TODO: Add spread operator as property that references a hidden spread function?

                let properties: any = {}

                let entity = 
                {
                    type: "Entity",
                    text: value.constructor.name,
                    properties: properties
                }

                for(var key of Object.getOwnPropertyNames(value))
                {
                    properties[key] = output(value[key])
                }

                return entity
            }   
        }
        else if(typeof(value) == "function")
        {
            return null
        }
        else
        {
            return null
        }
    }
}


export async function initialize(): Promise<void>
{
    Office.onReady(async () => 
    {

        Office.addin.setStartupBehavior(Office.StartupBehavior.load)

        // Disable cache.
        Office.context.document.settings.set("Office.ForceRefreshCustomFunctionsCache", true)  
        Office.context.document.settings.saveAsync()

        // Determine current document path, load the Webcellar root file and register exports.
        Office.context.document.getFilePropertiesAsync(null!, async (res) => 
        {
            if(res && res.value && res.value.url)
            {
                var excelFile = res.value.url
                var webcellarFile = ""

                if(webcellarFile == "")
                {
                    try
                    {
                        var resp = await fetch(`/${res.value.url}.js`, {method: "HEAD"})

                        if(resp.ok)
                        {
                            webcellarFile = `${res.value.url}.js`
                        }
                    }
                    catch {}
                }

                if(webcellarFile == "")
                {
                    try
                    {
                        var resp = await fetch(`/${res.value.url}.ts`, {method: "HEAD"})

                        if(resp.ok)
                        {
                            webcellarFile = `${res.value.url}.ts`
                        }
                    }
                    catch {}
                }

                document.body.innerHTML += `<p style="text-align: center; overflow-wrap: anywhere;"><strong>Current Excel file</strong>: ${excelFile.replaceAll("\\", "/")}</p>`
                document.body.innerHTML += `<p style="text-align: center; overflow-wrap: anywhere;"><strong>Current Webcellar file</strong>: ${webcellarFile.replaceAll("\\", "/")}</p>`

                if(webcellarFile == "")
                {

                }
                else
                {
                    try
                    {
                        var root = await importShim(`/${webcellarFile}`)
                        await register(root)
                    }
                    catch(e)
                    {
                        if(e instanceof Error)
                        {
                            var table = ""
                            table += `<table>`
                            table += `<tbody style="overflow-wrap: anywhere;">`
                            table += `<tr><th style="width: 40%">Key</th><th style="width: 60%">Val</th></tr>`

                            for(var key of new Set(["name", "message", "cause", "stack", ...Object.getOwnPropertyNames(e)]))
                            {
                                var val = ""
                                try
                                {
                                    val = (e as any)[key]
                                }
                                catch{}

                                table += `<tr><td>${key}</td><td>${val}</td></tr>`   
                            }

                            table += `</tbody>`
                            table += `</table>`

                            document.body.innerHTML += `<p style="text-align: center; overflow-wrap: anywhere;">${e.toString()}</p>`
                            document.body.innerHTML += table
                            document.body.innerHTML += `<style>td, th {vertical-align: top; text-align: left;}</style>`

                        }
                        else
                        {
                            document.body.innerHTML += `<p style="text-align: center; overflow-wrap: anywhere;">Error: ${e}</p>`
                        }
                    }
                }
            }
            else
            {
                document.body.innerHTML += `<p style="text-align: center; overflow-wrap: anywhere;">Warning: Unsaved document. Save the document and refresh Webcellar.</p>`
            }
        })
    })
}


async function register(root: any)
{
    // Reference: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-json
    var functions = 
    {
        allowCustomDataForDataTypeAny: true, // false => entities from Excel to Webcellar are JSON strings? 
        allowErrorForDataTypeAny: false,
        functions: [] as any[]
    } 

    var kvps = Object.entries(root)

    for(var i = 0; i < kvps.length; i++)
    {
        // Default namespace is A.
        kvps[i][0] = "A." + kvps[i][0]   
    }


    // Register exports.
    while(kvps.length > 0)
    {
        let [key, value] = kvps.pop()!

        let meta = new Meta()
        
        Object.assign(meta, metadata.get(value))

        let id = meta.excel?.id ?? key.toUpperCase()

        let inputHandler = meta.input == "convert" ? input : (value: any, root = false) => value
        let outputHandler = meta.output == "convert" ? output : (value: any, root = false) => value

        // TODO: Check how things are registered.
        // Use root[key] instead of value, because this may be beneficial for hot reloading (use "recursive" keys)?
        // Maybe always store key and parent object (not just value)?
        // Handle other fields (e.g., description) when registering?

        if(value == null)
        {
        }
        else if(typeof(value) == "string")
        {
            CustomFunctions.associate(id, () => {return outputHandler(value, true)})

            functions.functions.push(
            {
                id: id,
                name: id,
                parameters: [],
                result: 
                {
                    dimensionality: "matrix"
                },
                ...meta?.excel,
            })
        }
        else if(typeof(value) == "number")
        {
            CustomFunctions.associate(id, () => {return outputHandler(value, true)})

            functions.functions.push(
            {
                id: id,
                name: id,
                parameters: [],
                result: 
                {
                    dimensionality: "matrix"
                },
                ...meta?.excel,
            })
        }
        else if(typeof(value) == "bigint")
        {
        }
        else if(typeof(value) == "boolean")
        {
            CustomFunctions.associate(id, () => {return outputHandler(value, true)})

            functions.functions.push(
            {
                id: id,
                name: id,
                parameters: [],
                result: 
                {
                    dimensionality: "matrix"
                },
                ...meta?.excel,
            })
        }
        else if(typeof(value) == "symbol")
        {
        }
        else if(typeof(value) == "undefined")
        {
        }
        else if(typeof(value) == "object")
        {
            // TODO: Handle other cases, e.g., Map

            if(Array.isArray(value))
            {
                CustomFunctions.associate(id, (index: number | undefined) => 
                {
                    if(index == null)
                    {
                        return outputHandler(value, true)
                    }
                    else
                    {
                        return outputHandler(value[index], true)
                    }
                })

                functions.functions.push(
                {
                    id: id,
                    name: id,
                    parameters: 
                    [
                        {name: "index", type: "number", optional: true}
                    ],
                    result: 
                    {
                        dimensionality: "matrix"
                    },
                    ...meta?.excel
                })
            }
            else
            {
                CustomFunctions.associate(id, () => {return outputHandler(value, true)})

                functions.functions.push(
                {
                    id: id,
                    name: id,
                    parameters: [],
                    result: {dimensionality: "matrix"},
                    ...meta?.excel
                })

                // TODO: Handle infinite loops (i.e., objects that reference themselves)?
                // Add properties for consideration to be registered.
                for(var k of Object.getOwnPropertyNames(value))
                {
                    var prop = (value as any)[k]
                    
                    if(typeof(prop) == "function")
                    {
                        // Bind this and make sure that toString is matches original, bc argument types are extracted from it
                        var fn = prop.bind(value)
                        fn.toString = () => prop.toString()

                        kvps.push([`${id}.${k}`, fn])
                    }
                    else
                    {
                        kvps.push([`${id}.${k}`, prop])
                    }
                }

                // Add prototype properties and functions.
                var proto = Object.getPrototypeOf(value)
                while (proto && proto !== Object.prototype) 
                {
                    for(var k of Object.getOwnPropertyNames(proto))
                    {
                        if(k != "constructor")
                        {
                            var prop = (value as any)[k]
                    
                            if(typeof(prop) == "function")
                            {
                                // Bind this and make sure that toString is matches original, bc argument types are extracted from it
                                var fn = prop.bind(value)
                                fn.toString = () => prop.toString()

                                kvps.push([`${id}.${k}`, fn])
                            }
                            else
                            {
                                kvps.push([`${id}.${k}`, prop])
                            }
                        }   
                    }

                    proto = Object.getPrototypeOf(proto)
                }
            }
        }
        else if(typeof(value) == "function")
        {
            // Function parameter types are added as comments into the function source
            var source = value.toString()

            var parametersConStart = source.indexOf(`/*${utils.con.fence}`)
            var parametersConEnd = source.indexOf(`${utils.con.fence}*/`)

            if(parametersConStart >= 0 && parametersConEnd >= 0)
            {
                var parametersCon = source.slice(parametersConStart + `/*${utils.con.fence}`.length, parametersConEnd)
                var parameters = utils.con.parse(parametersCon)

                CustomFunctions.associate(id, async (...args: any[]) => 
                {
                    try
                    {
                        // Skip last bc it is Excel invocation
                        for(var i = 0; i < args.length - 1; i++)
                        {
                            args[i] = inputHandler(args[i], true)   
                        }

                        // Handle function vs constructor
                        try
                        {
                            var val = await value(...args)

                            return outputHandler(val, true)  
                        }
                        catch(error: any)
                        {
                            // console.log(error)
                            if(error instanceof TypeError)
                            {
                                try
                                {
                                    var val = await new (value as any)(...args)

                                    return outputHandler(val, true) 
                                }
                                catch(ee: any)
                                {
                                    // TODO: Improve?
                                    return outputHandler(ee.toString(), true)
                                }
                            }
                            else
                            {
                                // TODO: Improve?
                                return outputHandler(error.toString(), true)
                            }
                        }
                    }
                    catch(error: any)
                    {
                        // TODO: Improve?
                        return outputHandler(error.toString(), true)
                    }
                })

                functions.functions.push(
                {
                    id: id,
                    name: id,
                    description: (meta?.excel?.parameters ?? parameters).map((p: any) => `${p.name}: ${p.description}`).join(", "),
                    parameters: parameters,
                    result: {dimensionality: "matrix"},
                    ...meta?.excel
                })
            }
            else
            {
                console.log(`Webcellar: Could not register ${id}.`)
            }
        }
        else
        {
        }
    }

    // Register functions with Excel. 
    // TODO: Update to a documented solution (this is not officially documented in Office.js, but the ScriptLab add-in uses it).
    
    await (Excel.CustomFunctionManager as any).register(JSON.stringify(functions), "")
}

