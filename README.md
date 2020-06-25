# google sheets get cell data
 To make sheets avaiable for requesting follow official google guid
 https://developers.google.com/sheets/api/quickstart/js#python-3.x

 Don't forget to put your `CLIENT_ID` (line 20), `API_KEY` (line 21) and `SHEET_ID` (line 29) inside `index.html`.

 After you authorize you can use getCellData function in the console.
 Notice that the function is asyncronus and returns `Promise`, so to take a look on the result use `await` before function invocation or `then` of Promise object;

 # Function arguments
The function takes a string that can contain just a position like `a1`, `a 1`, `1a`, `1 a`. Character case doesn't matter. Or it can be like `row 1 column 2`, `row 1 2 column`, `row 1 column b`. Important thing is the number of a column or a row should go next to its key word (key word is `column` and `row`) and the string only should contain `column` and `row` and its positions. Key words can be in russian.
