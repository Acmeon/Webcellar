import * as webcellar from "/.webcellar/webcellar.js"

export var string = "Hello World!"

export var array = [1, 2, 3, 4]

export var matrix = [[1, 2], [3, 4]]

export var object = {a: 1, b: [2, 3], c: {c0: 4, c1: 5}, e: "six"}

export var add = (a: number, b: number) => a + b


export var number = 1234
export var boolean = true
export var undefinedx = undefined
export var nullx = null

export function concatenate(a: number[], b: number[])
{
    return [...a, ...b]
}

export async function delay(seconds: number)
{
    await new Promise(resolve => setTimeout(resolve, 1000 * seconds))

    return `Async function delayed for ${seconds} seconds.`
}

// State can be preserved, but should probably in most cases be avoided.
var state = 0.0
export function volatileStatefulRandomWalk()
{
    state += Math.random()
    return state
}
webcellar.meta(volatileStatefulRandomWalk, 
{
    excel:
    {
        // Reference: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-json#metadata-reference
        description: "This function recalculates when the Excel file is modified.",
        options: {volatile: true},
    }  
})


export function globalNamespace()
{
    return "Call this function from Excel with GLOBALNAMESPACE()."
}
webcellar.meta(globalNamespace, 
{
    // Reference: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-json#metadata-reference
    excel: 
    {
        id: "GLOBALNAMESPACE",
    }
})


export class Car
{
    brand: string
    speed: number

    constructor(brand: string, speed: number)
    {
        this.brand = brand
        this.speed = speed
    }

    str()
    {
        return `Brand: ${this.brand}, Speed: ${this.speed}`
    }
}

export var carInstance = new Car("Ferrari", 9000)



// import * as webcellar from "/.webcellar/webcellar.js"

export function timestamp(utc: boolean, invocation: CustomFunctions.StreamingInvocation<string>)
{
    var timer = setInterval(() => 
    {
        if(utc)
        {
            invocation.setResult(new Date().toUTCString())
        }
        else
        {
            invocation.setResult(new Date().toString())
        }

    }, 10000)

    invocation.onCanceled = () => 
    {
        clearInterval(timer)
    }
}

webcellar.meta(timestamp, 
{
    input: "raw",   // Default: "convert"
    output: "raw",  // Default: "convert"

    // Reference: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-json#functions
    excel: 
    {
        parameters: [{name: "utc", type: "boolean", dimensionality: "scalar"}],
        options: {stream: true},
        result: {dimensionality: "scalar"}
    }
})




import {sum} from "./module.ts"
export {sum}


export {foo} from "./foo/foo.xlsx.js"


export var test = [] as number[]


export var zzz = new class
{
    a = "Hello"
    b = "World!"
}

webcellar.meta(zzz, 
{
    excel: 
    {
        id: "ZZZ",
    }
})
