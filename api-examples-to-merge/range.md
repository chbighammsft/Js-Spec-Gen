### getHyperlinkRanges()
Word.run(function(context) {
    var hyperlinks = context.document.body.getRange().getHyperlinkRanges();
    hyperlinks.load();
    context.sync().then(function () {
        for (var i = 0; i < hyperlinks.items.length; i++) {
            var link = hyperlinks.items[];
            var mdLink = '[' + link.text +](' + link.hyperlink +') ';
            link.hyperlink = '';
            link.insertText(mdLink, 'Replace');
        }
    });
});