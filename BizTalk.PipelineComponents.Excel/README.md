# BizTalk.PipelineComponents.Excel
This is a Excel BizTalkpipeline Encoder and Decoder that uses the  [NPOI](https://www.nuget.org/packages/NPOI/) nuget package
for parsing and creation of Excel files.<p/>
The component uses exact position for Sheets, Rows and Cells controlled by XSD schemas.<p/>
For outgoing files i have opted to use a template file, really a regular Excel file formatted in the way you want.<br/>
This way you can concentrate on content and not formatting. 

### Workbook
The workbook is represented by the root node in the schema.

### Sheets
All record node bellow the root node represents Sheets. The Notes property in the schema property window specifies the exakt 0 based position of the sheet.

### Rows

Any record node bellow a Sheet node represents a row. The Notes property in the schema property window specifies the exakt 0 based position of the row in the specified sheet.<br/>
For the row, position, is more of a startposition as a row can exist multiple times, specified by the maxOccurs property. In this case Notes specifies From position and maxOccurs specifies To position.<br/>
Any row record can contain maxOccurs unbounded, but for incomming files only the last row record with maxOccurs unbounded will be processed, any row record after the first row with maxOccurs unbounded will be ignored.<br/>
For outgoing files it is possible to use maxOccurs unbounded for several rows.<br/>
If you specify a row to start on the same position as another row, information will be overwritten for outging files.

### Cells
Any element or attribute node bellow a Row node represents a cell in the Excel file. The Notes property in the schema property window specifies the exakt 0 based position of the cell inside the Excel row.<br/>
Only bellow XSD elements are allowed<br/>
- String
- Boolean
- Date
- DateTime
- Time
- Double

All other types will be handled as String type.

Formulas will also be processed correctly.

### Common
You do not need to specify all objects in the Schema from the source file only the content you are interested in, everything else will be ignored.

<img src="Schema1.jpg"/>

### Envelope
It is possible to use an envelope schema. Which requires some specific Schema configuration.<br/>
Schema must be set to Envelope=Yes and a Sheet(s) must be set to ounbounded. Yes its still possible to use multiple Sheets.

<img src="envelope.jpg"/>

The last Row record with the settings minOccurs="0" and maxOccurs="1" in the Sheet will be handled as the repeating row. Only one row will be processed as repeating row.

Bellow Row points to the second row in the Sheet

<img src="envelope_row.jpg"/>

All other Row record(s) will be added to each repeating row found. Meaning the information from the other records will be copied and bundled to each each repeating row.

Using the shema bellow Company record would be appended to each Person record.

<img src="envelope_bundle.jpg"/>