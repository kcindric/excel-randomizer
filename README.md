# excel-randomizer

I had a task where I wanted to randomize two pairs of data (area and legth of an object) and output it to an Excel workbook.
I used [LinqToExcel](https://github.com/paulyoder/LinqToExcel) to query my dataset from an existing Excel table and then [ClosedXML](https://github.com/ClosedXML/ClosedXML) to write the randomized data to a new workbook.

I created a new `Random()` class so I can use it to randomly order my query using Linq.

While using [LinqToExcel](https://github.com/paulyoder/LinqToExcel) you first need to create an class with properties names that are matchin the Excel table headers. Check the documentation for workaround on this subject.

Before the `.Orderby()` method you need to add `.ToList()` while working with LinqToExcel. You can find more info on this subject [here](https://stackoverflow.com/questions/55223165/cant-random-orderby-while-using-linqtoexcel).
