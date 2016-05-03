### getById(id: number)  
Word.run(function (ctx) {
   var firstList = ctx.document.body.lists.getListById(listId);
   return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});