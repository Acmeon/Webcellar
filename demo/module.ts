export function sum(array: number[])
{
    var s = 0

    for(var v of array)
    {
        s += v
    }

    return s
}