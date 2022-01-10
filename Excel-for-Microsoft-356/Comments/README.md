# Comments (Threaded Comments) and Notes

## What are Comments and Notes

### Threaded Comments

Comments that Excel supports since Excel for Microsoft 365, and Threaded Comments work **ONLY** on Excel for Microsoft 365 (or wild guess, the version with online account and online cooperation).  
Works like a discussion or a Todo item conversation sticked to any cell.  
And threaded comments will fallback and become a joined note on desktop versions of Excel (there is an example below).

```text
[Threaded comment]

Your version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924

Comment:
    third
MultiLine
Reply:
    reply1
Reply:
    reply2
```

### Notes

Notes are comments in earlier versions (desktop versions) of Excel.  
Display as a "pure text" (but supports bold, italic, underlined, etc.) yellow background box with an arrow linked to a cell.

## How Excel stores Comments and Notes

### Notes

- `content` store on `./xl/comments${N}.xml`
- `layout` store on `./xl/drawings/vmlDrawing${N}.vml`
- relation `ref` store on `./xl/worksheets/_rels/sheet${I}.xml.rels`
- `sheet` data store on `./xl/worksheets/sheet{I}.xml`

#### how they link to each other

for example: note on cell A9.

1. content

   ```xml
   <!-- content.ref == A9 -->
   <comments>
     <commentlist>
       <comment ref="A9" authorId="3" shapeId="0" xr:uid="{94AD8496-EE48-4F40-BB22-1BE638458927}">
               <text>
                   <r>
                       <rPr>
                           <sz val="11"/>
                           <color rgb="FF000000"/>
                           <rFont val="Calibri"/>
                           <family val="2"/>
                       </rPr>
                       <t xml:space="preserve">Jonham Chen:
                       </t>
                   </r>
                   <r>
                       <rPr>
                           <sz val="11"/>
                           <color rgb="FF000000"/>
                           <rFont val="Calibri"/>
                           <family val="2"/>
                       </rPr>
                       <t xml:space="preserve">a note
                       </t>
                   </r>
                   <r>
                       <rPr>
                           <u/>
                           <sz val="11"/>
                           <color rgb="FF000000"/>
                           <rFont val="Calibri"/>
                           <family val="2"/>
                       </rPr>
                       <t xml:space="preserve">with </t>
                   </r>
                   <r>
                       <rPr>
                           <b/>
                           <sz val="11"/>
                           <color rgb="FF000000"/>
                           <rFont val="Calibri"/>
                           <family val="2"/>
                       </rPr>
                       <t xml:space="preserve">BOLD </t>
                   </r>
                   <r>
                       <rPr>
                           <i/>
                           <sz val="11"/>
                           <color rgb="FF000000"/>
                           <rFont val="Calibri"/>
                           <family val="2"/>
                       </rPr>
                       <t>Text</t>
                   </r>
               </text>
           </comment>
     </commentlist>
   </comments>
   ```

2. layout

   ```xml
   <!-- <x:Row>8</x:Row><x:Column>0</x:Column> ==> [0, 8] ==> A9 -->
   <xml xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:x="urn:schemas-microsoft-com:office:excel">
       <v:shape id="_x0000_s1026" type="#_x0000_t202" style='position:absolute;margin-left:251pt;margin-top:113pt;width:106pt;height:67pt;z-index:3;visibility:hidden' fillcolor="infoBackground" strokecolor="#217346" o:insetmode="auto">
         <v:fill color2="infoBackground"/>
         <v:shadow color="none [81]" obscured="t"/>
         <v:path o:connecttype="none"/>
         <v:textbox style='mso-direction-alt:auto'>
           <div style='text-align:left'></div>
         </v:textbox>
         <x:ClientData ObjectType="Note">
           <x:MoveWithCells/>
           <x:SizeWithCells/>
           <x:Anchor>
         1, 15, 7, 8, 3, 15, 12, 0</x:Anchor>
           <x:AutoFill>False</x:AutoFill>
           <x:Row>8</x:Row>
           <x:Column>0</x:Column>
         </x:ClientData>
       </v:shape>
     </xml>
   ```

3. ref

   ```xml
    <!-- ref -> content: comment1.xml, layout: vmlDrawing1.vml -->
    <Relationships
      xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <!-- ... -->
      <Relationship Id="rId2"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        Target="../comments1.xml"/>
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
        Target="../drawings/vmlDrawing1.vml"/>
   </Relationships>
   ```

4. sheet

   ```xml
   <!-- sheet.relationship, sheetData.row[r=9].c[r="A9"] -->
   <worksheet
     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
   >
     <sheetData>
      <!-- ... -->
      <row r="9" spans="1:1" x14ac:dyDescent="0.2">
        <c r="A9" t="s">
          <v>3</v>
        </c>
      </row>
   </sheetData>
   </worksheet>
   ```
