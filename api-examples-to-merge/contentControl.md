### getRange(rangeLocation: RangeLocation)
Word.run(function (context) {
    var body = context.document.body;

    var contentControls = body.contentControls;
    context.load(contentControls);
    return context.sync()
        .then(function () {
            var contentControl = contentControls.items[0];
                    
            var range = contentControl.getRange("Start");
            range.insertParagraph("Start of range", "Before");
        })
            
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
        console.log("Trace info: " + JSON.stringify(error.traceMessages));
    }
});

### getTextRanges(punctuationMarks: string[], trimSpacing: bool)
Word.run(function (context) {
    var body = context.document.body;

    var contentControls = body.contentControls;
    context.load(contentControls);
    return context.sync()
        .then(function () {
            var contentControl = contentControls.items[0];
            context.load(contentControl);
            return context.sync()
               .then(function () {
                    var ranges = contentControl.getTextRanges([".", "!", "?"], true);
                    context.load(ranges);
                    return context.sync()
                        .then(function () {
                            contentControl.clear();
                            for (var range in ranges.items) {
                                contentControl.insertParagraph(range);
                            }
                        });
                })
        })
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
        console.log("Trace info: " + JSON.stringify(error.traceMessages));
    }
});