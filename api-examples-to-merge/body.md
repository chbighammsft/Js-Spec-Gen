### insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])
Word.run(function (context) {
    var body = context.document.body;
    context.load(body);
    return context.sync()
        .then(function () {
            body.insertTable(2, 2, "End",
                [
                    ["Column 1", "Column 2"],
                    ["Column Entry 1", "Column Entry 2"]
                ]);
        })
        .then(context.sync)
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});

### body.getRange(rangeLocation: RangeLocation)
Word.run(function (context) {
    var body = context.document.body;
    context.load(body);

    var range = body.getRange("Whole");

    range.select("End");
            
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
       console.log("Debug info: " + JSON.stringify(error.debugInfo));
       console.log("Trace info: " + JSON.stringify(error.traceMessages));
    }
});