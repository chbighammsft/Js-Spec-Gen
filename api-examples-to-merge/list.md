### insertParagraph(paragraphText: string, insertLocation: InsertLocation)  
Word.run(function (context) {
    var lists = context.document.body.lists;
    context.load(lists);
    return context.sync()
       .then(function () {
            var list = lists.items[listId];
            list.insertParagraph("Paragraph text", "Start");
        })
        .then(context.sync)
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});