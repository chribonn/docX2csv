# docX2csv


If this option where to have a place on the office toolbar it would be under the References Menu because, like many of the options there, it indexes a document and generates a summary of that index. Think about this option as the Table of Contents (ToC) option on steroids. 

*docX2csv* creates a csv file from a *docx* of the text associated with user-selected styles (more than one style can be selected) together with the section, page and line this text was on. If the optional **Header Style** is chosen the text written in it will be included in the csv file. 

The generated *csv* can be imported into a spreadsheet and manipulated to produce a summary of the sought information. 

### Why?

A use case of *docX2csv* would be to create a summary of the RACI matrix, also known as responsibility assignment matrix (RAM) or LRC (linear responsibility chart).  The document contains sections (or pages) describing each action. 

The scenario is that a document could consists of sections starting with a Chapter Heading, followed by text to identify who is Responsible, Accountable, Consulted and Informed. The document would contain additional text and images describing the topic related to this chapter.

The Chapter Heading, and each RACI element would be associated with a particular style. The styles could be visually identical because *docX2csv* processes on the style name.

Using *docX2csv* one can generate a summary of the actions and the RACI groups or individuals.




When the source document is updated it would be reprocessed and the resulting table would be updated.

### How

*docX2csv* creates the csv based on Styles in the document. The Title would be associated with a Style to the title (as happens with the TOC option in Word) and would select one or more additional styles that will be crossed reference. For example using the RACI case, each role could have its own style.  The engine automatically adds the section, page and page paragraph number to allow for easy cross referencing.  Being in columnar format, data that is not required can be easily ignored.


### Operation

Text that will be used in the index needs to be properly assigned a Style.  The formatting of the Style does not impact the performance of *docX2csv*. This means that different Styles may be visually identical and changing the format of the Style will not alter the operation of *docX2csv*.

## Conclusion

Every document that needs to be summarised is a candidate for *docX2csv*. The most efficient route is to plan beforehand what content is to be summarised and define and apply these styles consistently. It would be a good idea with long document to use colour coding for styles while the document is being developed (changing them to default black) at project end.


## Limitations

https://dadoverflow.com/2022/01/30/parsing-word-documents-with-python/