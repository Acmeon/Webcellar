// Comment Object Notation
export var con = 
{
    fence: "3dc43160-e0db-4678-8d54-ef9fd7eeb889",
    stringify(value: any)
    {
        var str = JSON.stringify(value)
                       .replace(/\*\//g, "*\\/")
                       .replace(/\r/g, "\\r")
                       .replace(/\n/g, "\\n")

        return str
    },
    parse(comment: string)
    {
        var str = comment.replace(/\*\\/g, "*/")
                         .replace(/\\r/g, "\r")
                         .replace(/\\n/g, "\n")

        return JSON.parse(str)
    }
}