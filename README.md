# google-drive-to-spreadsheet
Apps Script files for creating an incremental data load from Google Drive into Google Sheets


Template:
```
Search File	Folder	Size	<style>	<style>
"<div onclick=""ga('send', 'event', '{{URL}}');window.open('{{URL}}')"">
        <div class=""title"">
                <img class=""typeimg"" src=""{{Image}}"" />
        </div>
        <div style=""vertical-align:middle;"">
                <div class=""title"">{{Title}}<div>
                <div class=""description"">{{Description}}</div>
        </div>
</div>

<object hidden> {{Id}}

<object hidden> {{Labels}}

<object hidden> {{Folder}}
"	"<div onclick=""ga('send', 'event', '{{URL}}'); window.open('{{URL}}')"">
  <div style="" text-align: center;"">{{Folder}}</div>
</div>"	"<div onclick=""ga('send', 'event', '{{URL}}');window.open('{{URL}}')"">
  <div class=""size"">{{File Size}}</div>
</div>"	".typeimg {
width:80px;
height: 80px;
float:left;
}
body, td, div, span, p {
font-family:'Roboto', sans-serif;
color: #444444;
font-size: 15px;
}
.title {
  color: #004f83;
  font-size: 18px;
text-decoration: none;
margin-left: -3px!important;
}
.description {
text-align:justify;
}
.owner {
color:#999999;
}
tr {
cursor:hand;
}
.size {
width:100%;
text-align:center;
padding:5px!important;
}



<div class=""awt-searchFilter-cont"">
        <input class=""awt-searchFilter-input"" placeholder=""Search For File"" aria-label=""Search Filter. Placeholder:Search For File"" tabindex=""0"">
</div>

"	"/*** Awesome Table ***/

/* filtered results information : last - first / total number of result */
.count {
        border-top:1px solid #E5E5E5;
        border-bottom:1px solid #E5E5E5;
}

/*** Google visualization override ***/

/* page number display */
.google-visualization-table-page-numbers {
        display:none !important;
}

/* Table view header */
.google-visualization-table-table .gradient,
.google-visualization-table-div-page .gradient {
        background: #3C81F8 !important
        color:#ffffff;
}

/* selected/hovered row in a TABLE view */
.google-visualization-table-tr-sel td,
.google-visualization-table-tr-over td {
        background-color: #D6E9F8!important;
}

/*** Configuration of FILTERS ***/

/** Labels of filters **/
.google-visualization-controls-label {
        color:#333;
}

/** StringFilter **/
.google-visualization-controls-stringfilter INPUT {
        border:1px solid #d9d9d9!important;
        color:#222;
}
.google-visualization-controls-stringfilter INPUT:hover {
        border:1px solid #b9b9b9;
        border-top:1px solid #a0a0a0;
}
.google-visualization-controls-stringfilter INPUT:focus {
        border:1px solid #4d90fe;
}

/** CategoryFilter **/
.google-visualization-controls-categoryfilter .charts-menu-button,
.google-visualization-controls-categoryfilter .charts-menu-button.charts-menu-button-hover,
.google-visualization-controls-categoryfilter .charts-menu-button.charts-menu-button-active {
        color:#444;
        border:1px solid rgba(0,0,0,0.1);
        background:none;
        background:#f5f5f5;
}
.google-visualization-controls-categoryfilter LI{
        background-color:#3C81F8!important;
        color:#FFF;
        height:25px;
        vertical-align: middle;
        border-radius: 5px;
        line-height:25px;

        }

/* CategoryFilter & csvFilter hovered item in the dropdown */
.charts-menuitem-highlight {
        background-color:#437AF8!important;
        border-color:#3C81F8!important;
        color:#FFF;
}


/* header */
.google-visualization-table-table .gradient, .google-visualization-table-div-page .gradient {
background: #004f83 !important;
  color: #FFFFFF;
  text-transform: uppercase;
  font-weight: normal;
}

.google-visualization-controls-categoryfilter LI {
  background-color: #437AF8!important;
}

.google-visualization-controls-categoryfilter .charts-menu-inner-box {
background-color: #437AF8!important;
}

.google-visualization-controls-label {
font-weight: 500!important;        
}

.charts-menuitem-highlight>.charts-menuitem-content {
color: white;
}

"
```
