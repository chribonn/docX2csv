# docX2csv: Microsoft Word Style Summary Creator

![](../assets/docX2csv.png)

## Introduction

### What is docX2csv?

*docX2csv* is a tool that takes a docx document and creates an indexed summary of the content. If this were a standard feature of Microsoft Word, you would find it in the references menu next to the Table of Contents (ToC) option under the Reference menu. You can think of *docX2csv* as the table of contents option on steroids.

![](../assets/docX2csv_001_1708519015139_0.png)

### How did it come about?

*docX2csv* was developed to help with the generation of a RACI matrix from a long and constantly changing document. Multiple people were contributing to this document, and the matrix ended up being out of sync constantly.

Having to manually work through the document to update this matrix was very time-consuming and error-prone. On the other hand, the ToC was a breeze using the built-in option. With the ToC, the only action was to right click on the existing table of contents and rebuilt it.

Just like the ToC option, *docX2csv* uses styles to generate its output.

### What are the use cases?

The use cases for this utility are in compliance, project management, security, and any sphere where a document needs to be summarised. Besides RACI matrixes, it can be used to generate SWOT, Risk Quadrant, and Competitive analysis matrices. There will be other needs cases not envisaged today.

## Running docX2csv

![](../assets/docX2csv_1707386751208_0.gif)

When you open *docX2csv*, you are presented with a GUI. You click on the **Browse** button and choose the Word document you want to process.

*docX2csv* reads the document and populates the **CrossRef Styles** list box and **Header Style** drop down with the styles used in your document.

**CrossRef Styles** is mandatory. You can multiselect styles. For example, you have a RACI document, and each of the four categories of Responsible, Accountable, Consulted, and Informed have their own style. The styles selected in this listbox will act as the index, meaning that every time a style is matched, it will be noted together with the document section, page, and paragraph number.

**Headers Style** is optional. If you select a header style, *docX2csv* will cross reference **CrossRef Styles** against the **Headers Style**. Cross-referencing can be done on document sections or pages.

The **Output Filename** is a read-only box showing where the generated CSV will be placed.

The *Run* button becomes enabled and will process the file, and generate the CSV output.

**docX2csv** automatically terminates on completion.

The resulting CSV file can then be opened in a spreadsheet and processed further to generate the desired summaries.

## Case Studies

A sample document called **SWOT.docx** is used to explain the operation of *docX2csv*. A series of walk-throughs will be performed on this document.

![](../assets/docX2csv-002_1707392123517_0.png)

The document can be download from this link: SWOT.docx. The styles within SWOT.docx contain the Section and Page number they are on. This has been done to enable one to follow along and better understand the output.

You may want to trigger the Show/Hide option on the toolbar and display the Section the cursor is on in Word to better follow some of the functionality of docX2csv.

### Case Study #01

In this scenario the Header Style is not selected.

#### Processing Screen

![](../assets/docX2csv-001_1707390864512_0.png)

#### Generated csv

| Style      | Style Text                        | Header Style Text | Section | Page | Paragraph | Linked Ref |
|------------|-----------------------------------|-------------------|---------|------|-----------|------------|
| Weaknesses | Weakness (Page 1 / Section 1)     | 1                 | 1       | 8    |           |            |
| Weaknesses | Weakness (Page 2 / Section 1)     | 1                 | 2       | 1    |           |            |
| Threats    | Threats Text (Page 2 / Section 2) | 2                 | 2       | 14   |           |            |
| Weaknesses | Weakness (Page 5 / Section 3)     | 3                 | 5       | 5    |           |            |

The generated csv file consists of the following columns (empty columns will be skipped):

* **Style**: This is the Style name that was matched
* **Style Text**: This is the text assigned to the matched Style
* **Section**: The section in the document where the Style was used
* **Page**: The page in the document where the Style was used
* **Paragraph**: The paragraph marks where the Style was used.

The style Weakness was used in 3 places in the document and was associated with the text shown under Style Text.

*docX2csv* processes styles based on the document's style name and not their formatting. In the SWOT.docx document these are coloured to allow one to follow more easily but they can be identically formatted if this is the formatting chosen for the source document.

### Case Study #02

In this scenario, the CrossRef Styles are Threats and Weaknesses just like Case Study #01 and we add the Header Style of Heading1. We Align on Section.

#### Processing Screen

![](../assets/docX2csv-003_1707394954188_0.png)

#### Generated csv

| Style      | Style Text                        | Header Style Text             | Section | Page | Paragraph | Linked Ref                |
|------------|-----------------------------------|-------------------------------|---------|------|-----------|---------------------------|
| Weaknesses | Weakness (Page 1 / Section 1)     | Header 1 (Page 1 / Section 1) | 1       | 1    | 8         | Weaknesses263709808978788 |
| Weaknesses | Weakness (Page 1 / Section 1)     | Header 2 (Page 2 / Section 1) | 1       | 1    | 8         | Weaknesses263709808978788 |
| Weaknesses | Weakness (Page 2 / Section 1)     | Header 1 (Page 1 / Section 1) | 1       | 2    | 1         | Weaknesses107532737854584 |
| Weaknesses | Weakness (Page 2 / Section 1)     | Header 2 (Page 2 / Section 1) | 1       | 2    | 1         | Weaknesses107532737854584 |
| Threats    | Threats Text (Page 2 / Section 2) | Header 3 (Page 3 / Section 2) | 2       | 2    | 14        | Threats182368771509297    |
| Threats    | Threats Text (Page 2 / Section 2) | Header 4 (Page 4 / Section 2) | 2       | 2    | 14        | Threats182368771509297    |
| Weaknesses | Weakness (Page 5 / Section 3)     | Header 5 (Page 5 / Section 3) | 3       | 5    | 5         | Weaknesses190353366798    |

In this scenario, whenever a cross referenced style is matched, we also include any header styles that are within the same section. If we take Threats, this is found in Section 2, Page 2, and Paragraph 14. Since we are aligning on sections, within this section there are two occurrences of Header Styles Heading1. Each combination generates its own row, thereby resulting in two rows for this threat.

The column Linked Ref is identical for match rows. For example, Threats both have the same reference of Weaknesses263709808978788. This allows one to perform actions on unique matches.

Weakness of Section 3, Page 5, Paragraph 5 was matched with only one **Header Style**, hence there is only one Linked Ref.

### Case Study #03

#### Processing Screen

Here we search CrossRef StyleStrengths including the Header Style Heading1 and Align on Pages.

![](../assets/docX2csv-004_1707396616010_0.png)

#### Generated csv

| Style     | Style Text                     | Header Style Text             | Section | Page | Paragraph | Linked Ref               |
|-----------|--------------------------------|-------------------------------|---------|------|-----------|--------------------------|
| Strengths | Strengths (Page 2 / Section 2) | Header 2 (Page 2 / Section 1) | 2       | 2    | 11        | Strengths240756528094835 |
| Strengths | Strength (Page 6 / Section 4)  | 4                             | 6       | 6    |           |                          |

The case identical to the first entry has already been discussed. The only difference is that since the alignment is on page rather than section, the only link would be with matched **Header Style** on the same page.

The second match is for the Strengths style on page 6, section 4. As there is no **Header Style** text on this page, a link did not happen. This is reflected by a blank **Linked Ref**. A blank **Header Style Text** does not mean that there was no header text; it could mean that the style was applied to a non-printable character such as a space or paragraph mark (see FAQ below).

### Processing the output

#### Processing Screen

![](../assets/docX2csv-005_1707400850059_0.png)

#### Generated csv

| Style         | Style Text                        | Header Style Text             | Section | Page | Paragraph | Linked Ref                   |
|---------------|-----------------------------------|-------------------------------|---------|------|-----------|------------------------------|
| Opportunities | Opportunity (Page 1 / Section 1)  | Header 1 (Page 1 / Section 1) | 1       | 1    | 3         | Opportunities195317518820605 |
| Opportunities | Opportunity (Page 1 / Section 1)  | Header 2 (Page 2 / Section 1) | 1       | 1    | 3         | Opportunities195317518820605 |
| Strengths     | Strengths (Page 2 / Section 2)    | Header 3 (Page 3 / Section 2) | 2       | 2    | 11        | Strengths94318597188847      |
| Strengths     | Strengths (Page 2 / Section 2)    | Header 4 (Page 4 / Section 2) | 2       | 2    | 11        | Strengths94318597188847      |
| Strengths     | Strength (Page 6 / Section 4)     | 4                             | 6       | 6    |           |                              |
| Threats       | Threats Text (Page 2 / Section 2) | Header 3 (Page 3 / Section 2) | 2       | 2    | 14        | Threats119391098062403       |
| Threats       | Threats Text (Page 2 / Section 2) | Header 4 (Page 4 / Section 2) | 2       | 2    | 14        | Threats119391098062403       |
| Weaknesses    | Weakness (Page 1 / Section 1)     | Header 1 (Page 1 / Section 1) | 1       | 1    | 8         | Weaknesses125310991226134    |
| Weaknesses    | Weakness (Page 1 / Section 1)     | Header 2 (Page 2 / Section 1) | 1       | 1    | 8         | Weaknesses125310991226134    |
| Weaknesses    | Weakness (Page 2 / Section 1)     | Header 1 (Page 1 / Section 1) | 1       | 2    | 1         | Weaknesses195254449828523    |
| Weaknesses    | Weakness (Page 2 / Section 1)     | Header 2 (Page 2 / Section 1) | 1       | 2    | 1         | Weaknesses195254449828523    |
| Weaknesses    | Weakness (Page 5 / Section 3)     | Header 5 (Page 5 / Section 3) | 3       | 5    | 5         | Weaknesses112118839087962    |

The data can be imported into a tool such as a spreadsheet, where it can be manipulated, filtered, summarised, and analysed further.

![](../assets/docX2csv2-002_1707408488989_0.gif)

## Know Limitations / FAQ

### Page numbers may be offset.

In Word, there are three categories of page breaks:

* Page Breaks (CTRL+ENTER or Menu: Layout -> Breaks -> (Page Break) Page)
* Section Page Breaks (Menu: Layout -> Breaks -> (Section Break) Next Page)
* Rendered Page Breaks
  
The first two are inserted by the user, while the third is generated by Word wherever it feels a page break should go. *Rendered Page Breaks* may be missing in the document.

Certain content is dynamically generated in the document on the fly and will not be factored into the page count. ToC entries are a case in point.

### Work Around

1. Insert page breaks rather than rely on rendered ones.
2. When there is an incorrect page identification, the page offset is consistent. This means that all pages that follow the one where the offset took place will need to be adjusted by the same amount . In your CSV simply add the offset to the page column.
   
### Why are there empty rows associated with a style or cross referenced against a header?

The most likely reason is that you have a blank paragraph associated with the style. User Word's Find option to search for them.

### Where can I download / fork / comment on the latest version:

https://github.com/chribonn/docX2csv
