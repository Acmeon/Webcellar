import * as acorn from "acorn"
import MagicString from "magic-string"
import {tsPlugin}  from "@sveltejs/acorn-typescript"
import * as utils from "./utils.ts"

var parser = acorn.Parser.extend(tsPlugin())

function walk(node: acorn.Node, callback: (node: acorn.Node) => void)
{
    // Custom walk function, because acorn-walk is unable to handle some TS nodes...

    // TODO: Check for correctness
    for(let [key, value] of Object.entries(node))
    {
        if(value != null && typeof(value) == "object" && "start" in value && "end" in value && "type" in value)
        {
            // Acorn Node
            callback(value)
            walk(value, callback)
        }
        else if(Array.isArray(value))
        {
            for(let v of value)
            {
                if(value != null && typeof(v) == "object" && "start" in v && "end" in v && "type" in v)
                {
                    // Acorn Node
                    callback(v)
                    walk(v, callback)
                }
            }   
        }
        else
        {
        }
    }
}

export function transform(source: string, url = "")
{
    var transformed = new MagicString(source)

    var ast = parser.parse(source, {ecmaVersion: "latest", sourceType: "module", locations: true})

    walk(ast, (node) =>
    {
        var n = node as acorn.AnyNode

        if(n.type == "FunctionDeclaration" || n.type == "FunctionExpression" || n.type == "ArrowFunctionExpression")
        {
            // Reference: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-json#parameters
            var parameters = []
            for(var p of n.params)
            {
                // TODO: Handle other cases.
                if(p.type == "Identifier")
                {
                    var typeAnnotation = (p as any).typeAnnotation as acorn.AnyNode

                    if(typeAnnotation)
                    {
                        var annotation = source.slice(typeAnnotation.start, typeAnnotation.end).replace(":", "").trim()
                        var annotationWithoutBrackets = annotation.replace(/(\[\])*$/, "").trim()

                        // TODO: Handle union of Excel "primitives" (number, string, boolean), e.g., number | boolean?
                        if(["any", "boolean", "number", "string"].includes(annotationWithoutBrackets))
                        {
                            parameters.push(
                            {
                                name: p.name,
                                description: annotation,
                                dimensionality: annotation.endsWith("[]") ? "matrix" : "scalar",
                                type: annotationWithoutBrackets,
                            })
                        }
                        else
                        {
                            parameters.push(
                            {
                                name: p.name,
                                description: annotation,
                                dimensionality: annotation.endsWith("[]") ? "matrix" : "scalar",
                                type: "any",
                            })
                        }
                    }
                    else
                    {
                        // Untyped case (allow all).
                        parameters.push(
                        {
                            name: p.name,
                            description: "any",
                            dimensionality: "matrix",
                            type: "any",
                        })
                    }
                }
            }

            // TODO: Improve the function parameter handling and avoid inserting type info into source? 
            // This approach ensures that the parameters can be accessed easily when registering functions in Excel, without considering references.

            var paramsAsComment = `/*${utils.con.fence}${utils.con.stringify(parameters)}${utils.con.fence}*/` 

            if(n.type == "FunctionDeclaration" || n.type == "FunctionExpression")
            {
                transformed.appendRight(n.body.start + 1, paramsAsComment)
            }
            else if(n.type == "ArrowFunctionExpression")
            {
                transformed.appendRight(n.body.start, paramsAsComment)
            }
        }

        // Remove TS annotations
        var nn = node as any
        if (nn.typeAnnotation) 
        {
            transformed.remove(nn.typeAnnotation.start, nn.typeAnnotation.end)
            // transformed.update(nn.typeAnnotation.start, nn.typeAnnotation.end, source.slice(nn.typeAnnotation.start, nn.typeAnnotation.end).replace(/[^\n]/g, ' '))
        }

        if (nn.type == "TSAsExpression") 
        {
            var lenTot = nn.end - nn.start 
            var lenExpr = nn.expression.end - nn.expression.start

            transformed.update(nn.start, nn.end, source.slice(nn.expression.start, nn.expression.end))
            // transformed.update(nn.start, nn.end, source.slice(nn.expression.start, nn.expression.end) + " ".repeat(lenTot - lenExpr))
        }

        if (nn.type == "TSInterfaceDeclaration" || nn.type == "TSTypeAliasDeclaration" || nn.type == "TSTypeParameterInstantiation") 
        {
            transformed.remove(nn.start, nn.end)
            // transformed.update(nn.start, nn.end, source.slice(nn.start, nn.end).replace(/[^\n]/g, ' '))
        }
    })

    // TODO: Verify that this is the best approach.
    var map = transformed.generateMap(
    {
        file: `${url}.map`,
        source: url,
        includeContent: true
    })

    transformed.append(`\n//# sourceMappingURL=${map.toUrl()}`)
    transformed.append(`\n//# sourceURL=${url}`)

    return transformed.toString()
}


if(typeof(window) !== "undefined")
{
    window.esmsInitOptions = 
    {
        shimMode: true,
        fetch: async (url, options) => 
        {
            var res = await fetch(url, options)

            if(!res.ok)
            {
                return res   
            }

            if(res.url.endsWith(".ts"))
            {
                var source = await res.text()
                var transformed = transform(source)

                return new Response(new Blob([transformed], {type: "application/javascript"}))
            }
            else
            {
                return res
            }
        }
    }
}
