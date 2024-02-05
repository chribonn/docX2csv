# docX2csv

If this option where to have a place on the office toolbar it would be under the References Menu because, like many of the options there, it indexes a document and generates a summary of that index.

*docX2csv* creates a csv file based on cross referencing user-selected styles [multiple may be selected] against another Style, section, page, paragraph mark count of the document.  The resulting csv can be imported into a csv manipulation tool (eg spreadsheet) where additional actions can be performed.

### Why?

A use case of *docX2csv* would be to create a summary of the RACI matrix, also known as responsibility assignment matrix (RAM) or LRC (linear responsibility chart).  The document contains sections (or pages) describing each action. Each action consists of an action title, description text as well as entries describing who is Responsible, Accountable, Consulted and Informed.  

Using *docX2csv* one can generate a summary of the actions and the RACI groups or individuals.




When the source document is updated it would be reprocessed and the resulting table would be updated.

### How

*docX2csv* creates the csv based on Styles in the document. The Title would be associated with a Style to the title (as happens with the TOC option in Word) and would select one or more additional styles that will be crossed reference. For example using the RACI case, each role could have its own style.  The engine automatically adds the section, page and page paragraph number to allow for easy cross referencing.  Being in columnar format, data that is not required can be easily ignored.


### Operation

Text that will be used in the index needs to be properly assigned a Style.  The formatting of the Style does not impact the performance of *docX2csv*. This means that visually different Styles may be visually the same and changing the format of the Style will not alter the operation of *docX2csv*.

## Conclusion

Every document that needs to be summarised is a candidate for *docX2csv*. The most efficient route is to plan beforehand what content is to be summarised and define and apply these styles consistently. It would be a good idea with long document to use colour coding for styles while the document is being developed (changing them to default black) at project end.
