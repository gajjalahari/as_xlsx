create or replace package as_xlsx
is
/**********************************************
**
** Author: Anton Scheffer
** Date: 19-02-2011
** See also: http://technology.amis.nl/blog/?p=10995
**
** Changelog:
**   Date: 21-02-2011
**     Added Aligment, horizontal, vertical, wrapText
**   Date: 06-03-2011
**     Added Comments, MergeCells, fixed bug for dependency on NLS-settings
**   Date: 16-03-2011
**     Added bold and italic fonts
**   Date: 22-03-2011
**     Fixed issue with timezone's set to a region(name) instead of a offset
**   Date: 08-04-2011
**     Fixed issue with XML-escaping from text
**   Date: 27-05-2011
**     Added MIT-license
**   Date: 11-08-2011
**     Fixed NLS-issue with column width
**   Date: 29-09-2011
**     Added font color
**   Date: 16-10-2011
**     fixed bug in add_string
**   Date: 26-04-2012
**     Fixed set_autofilter (only one autofilter per sheet, added _xlnm._FilterDatabase)
**     Added list_validation = drop-down
**   Date: 27-08-2013
**     Added freeze_pane
**   Date: 05-09-2013
**     Performance
**   Date: 14-07-2014
**      Added p_UseXf to query2sheet
**   Date: 23-10-2014
**      Added xml:space="preserve"
**   Date: 29-02-2016
**     Fixed issue with alignment in get_XfId
**     Thank you Bertrand Gouraud
**   Date: 01-04-2017
**     Added p_height to set_row
**   Date: 01-09-2017
**     Added add_image
**   Date: 06-09-2017
**     set_column_width: round column width to pixel
**   Date: 21-09-2017
**     added  c_version
**   Date: 29-09-2017
**     added  location and tooltip to hyperlink
**   Date: 23-05-2018
**     fixed bug in add_string (thank you David Short)
**     added tabColor to new_sheet
**   Date: 13-06-2018
**     added formulas
**   Date: 12-02-2020
**     added sys_refcursor overload of query2sheet
**     use default date format in query2sheet
**     changed to date1904=false
**   Date: 05-10-2023
**     added as_read_xlsx
**   Date: 15-11-2023
**     Support for reading Strict Open XML Spreadsheet
**     Added clob_val and string_len
**   Date: 18-10-2024
**     fix for add_image with images larger than 2000 bytes
**   Date: 21-11-2024
**     fix for dates before March 1900
**     additions to query2sheet
**   Date: 03-12-2024
**     Some display properties for a sheet
**   Date: 08-12-2024
**     Add encryption
**     Add null cells option to read function
**     Added tables
**     Fixed issues with using combinations
**       of comments, images, hyperlinks and tables on multiple sheets
**   Date: 27-01-2025
**      fixed issue with get_XfId
**   Date: 01-02-2025
**      fixed BUG with procedure cell with varchar2 value
**   Date: 17-06-2025
**      fixed BUG with read dates
**   Date: 22-06-2025
**     added formulas again
******************************************************************************
******************************************************************************
Copyright (C) 2011, 2025 by Anton Scheffer

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

******************************************************************************
******************************************** */
--
  use_utl_file    constant boolean := true;
  use_dbms_crypto constant boolean := false;
  --
  type tp_alignment is record
    ( vertical varchar2(11)
    , horizontal varchar2(16)
    , wrapText boolean
    );
  type sheet_names is table of varchar2(4000);
  type tp_one_cell is record
    ( sheet_nr number(2)
    , sheet_name varchar(4000)
    , row_nr number(10)
    , col_nr number(10)
    , cell varchar2(100)
    , cell_type varchar2(1)
    , string_val varchar2(4000)
    , number_val number
    , date_val date
    , formula varchar2(4000)
    , clob_val clob
    , string_len integer
  );
  type tp_all_cells is table of tp_one_cell;
--
  procedure clear_workbook;
--
  procedure new_sheet
    ( p_sheetname      varchar2    := null
    , p_tabcolor       varchar2    := null -- this is a hex ALPHA Red Green Blue value
    , p_show_gridlines boolean     := null
    , p_grid_color_idx pls_integer := null -- index in default color palette 0 - 65
    , p_show_headers   boolean     := null
    );
--
  function OraFmt2Excel( p_format varchar2 := null )
  return varchar2;
--
  function get_numFmt( p_format varchar2 := null )
  return pls_integer;
--
  function get_font
    ( p_name varchar2
    , p_family pls_integer := 2
    , p_fontsize number := 11
    , p_theme pls_integer := 1
    , p_underline boolean := false
    , p_italic boolean := false
    , p_bold boolean := false
    , p_rgb varchar2 := null -- this is a hex ALPHA Red Green Blue value
    )
  return pls_integer;
--
  function get_fill
    ( p_patternType varchar2
    , p_fgRGB varchar2 := null -- this is a hex ALPHA Red Green Blue value
    )
/*
none
solid
darkGray
mediumGray
lightGray
gray125
gray0625
darkHorizontal
darkVertical
darkDown
darkUp
darkGrid
darkTrellis
lightHorizontal
lightVertical
lightDown
lightUp
lightGrid
lightTrellis
*/
  return pls_integer;
--
  function get_border
    ( p_top    varchar2 := 'thin'
    , p_bottom varchar2 := 'thin'
    , p_left   varchar2 := 'thin'
    , p_right  varchar2 := 'thin'
    , p_rgb    varchar2 := null -- this is a hex ALPHA Red Green Blue value
    )
/*
none
thin
medium
dashed
dotted
thick
double
hair
mediumDashed
dashDot
mediumDashDot
dashDotDot
mediumDashDotDot
slantDashDot
*/
  return pls_integer;
--
  function get_alignment
    ( p_vertical varchar2 := null
    , p_horizontal varchar2 := null
    , p_wrapText boolean := null
    )
/* horizontal
center
centerContinuous
distributed
fill
general
justify
left
right
*/
/* vertical
bottom
center
distributed
justify
top
*/
  return tp_alignment;
  --
  function get_xfid
    ( p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    )
  return pls_integer;
  --
  procedure cell
    ( p_col   pls_integer
    , p_row   pls_integer
    , p_value number
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    );
--
  procedure cell
    ( p_col   pls_integer
    , p_row   pls_integer
    , p_value varchar2
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    );
--
  procedure cell
    ( p_col   pls_integer
    , p_row   pls_integer
    , p_value date
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    );
--
  procedure hyperlink
    ( p_col pls_integer
    , p_row pls_integer
    , p_url varchar2 := null
    , p_value varchar2 := null
    , p_sheet pls_integer := null
    , p_location varchar2 := null
    , p_tooltip varchar2 := null
    );
  --
  procedure num_formula
    ( p_col           pls_integer
    , p_row           pls_integer
    , p_formula       varchar2
    , p_default_value number := null
    , p_numFmtId      pls_integer  := null
    , p_fontId        pls_integer  := null
    , p_fillId        pls_integer  := null
    , p_borderId      pls_integer  := null
    , p_alignment     tp_alignment := null
    , p_sheet         pls_integer  := null
    , p_xfId          pls_integer  := null
    );
  --
  procedure str_formula
    ( p_col           pls_integer
    , p_row           pls_integer
    , p_formula       varchar2
    , p_default_value varchar2 := null
    , p_numFmtId      pls_integer  := null
    , p_fontId        pls_integer  := null
    , p_fillId        pls_integer  := null
    , p_borderId      pls_integer  := null
    , p_alignment     tp_alignment := null
    , p_sheet         pls_integer  := null
    , p_xfId          pls_integer  := null
    );
  --
  procedure date_formula
    ( p_col           pls_integer
    , p_row           pls_integer
    , p_formula       varchar2
    , p_default_value date := null
    , p_numFmtId      pls_integer  := null
    , p_fontId        pls_integer  := null
    , p_fillId        pls_integer  := null
    , p_borderId      pls_integer  := null
    , p_alignment     tp_alignment := null
    , p_sheet         pls_integer  := null
    , p_xfId          pls_integer  := null
    );
  --
  procedure comment
    ( p_col pls_integer
    , p_row pls_integer
    , p_text varchar2
    , p_author varchar2 := null
    , p_width pls_integer := 150  -- pixels
    , p_height pls_integer := 100  -- pixels
    , p_sheet pls_integer := null
    );
--
  procedure mergecells
    ( p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_sheet pls_integer := null
    );
--
  procedure list_validation
    ( p_sqref_col pls_integer
    , p_sqref_row pls_integer
    , p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_style varchar2 := 'stop' -- stop, warning, information
    , p_title varchar2 := null
    , p_prompt varchar := null
    , p_show_error boolean := false
    , p_error_title varchar2 := null
    , p_error_txt varchar2 := null
    , p_sheet pls_integer := null
    );
--
  procedure list_validation
    ( p_sqref_col pls_integer
    , p_sqref_row pls_integer
    , p_defined_name varchar2
    , p_style varchar2 := 'stop' -- stop, warning, information
    , p_title varchar2 := null
    , p_prompt varchar := null
    , p_show_error boolean := false
    , p_error_title varchar2 := null
    , p_error_txt varchar2 := null
    , p_sheet pls_integer := null
    );
--
  procedure defined_name
    ( p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_name varchar2
    , p_sheet pls_integer := null
    , p_localsheet pls_integer := null
    );
--
  procedure set_column_width
    ( p_col pls_integer
    , p_width number
    , p_sheet pls_integer := null
    );
--
  procedure set_column
    ( p_col pls_integer
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    );
--
  procedure set_row
    ( p_row pls_integer
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    , p_height number := null
    );
--
  procedure freeze_rows
    ( p_nr_rows pls_integer := 1
    , p_sheet pls_integer := null
    );
--
  procedure freeze_cols
    ( p_nr_cols pls_integer := 1
    , p_sheet pls_integer := null
    );
--
  procedure freeze_pane
    ( p_col pls_integer
    , p_row pls_integer
    , p_sheet pls_integer := null
    );
--
  procedure set_autofilter
    ( p_column_start pls_integer := null
    , p_column_end pls_integer := null
    , p_row_start pls_integer := null
    , p_row_end pls_integer := null
    , p_sheet pls_integer := null
    );
--
  procedure set_table
    ( p_column_start pls_integer
    , p_column_end   pls_integer
    , p_row_start    pls_integer
    , p_row_end      pls_integer
    , p_style        varchar2
    , p_name         varchar2    := null
    , p_sheet        pls_integer := null
    );
/* p_style
  TableStyleLight1  - TableStyleLight21
  TableStyleMedium1 - TableStyleMedium28
  TableStyleDark1   - TableStyleDark11
*/
--
  procedure set_tabcolor
    ( p_tabcolor varchar2 -- this is a hex ALPHA Red Green Blue value
    , p_sheet pls_integer := null
    );
--
  function finish( p_password varchar2  := null )
  return blob;
--
  procedure save
    ( p_directory varchar2
    , p_filename varchar2
    , p_password varchar2 := null
    );
  --
  procedure query2sheet
    ( p_sql            varchar2
    , p_column_headers boolean     := true
    , p_directory      varchar2    := null
    , p_filename       varchar2    := null
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    );
  --
  procedure query2sheet
    ( p_rc             in out sys_refcursor
    , p_column_headers boolean     := true
    , p_directory      varchar2    := null
    , p_filename       varchar2    := null
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    );
  --
  function query2sheet
    ( p_sql            varchar2
    , p_column_headers boolean     := true
    , p_directory      varchar2    := null
    , p_filename       varchar2    := null
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    )
  return number;  -- number of records returned by the query
  --
  function query2sheet
    ( p_rc             in out sys_refcursor
    , p_column_headers boolean     := true
    , p_directory      varchar2    := null
    , p_filename       varchar2    := null
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    )
  return number;  -- number of records returned by the query
  --
  procedure setUseXf( p_val boolean := true );
--
/*
use p_width and p_height to pass the size of an image which is not a png, jpg, gif or bmp
and even then I'm not sure if it will display an image :)
*/
  procedure add_image
    ( p_col pls_integer
    , p_row pls_integer
    , p_img blob
    , p_name varchar2 := ''
    , p_title varchar2 := ''
    , p_description varchar2 := ''
    , p_scale number := null
    , p_sheet pls_integer := null
    , p_width pls_integer := null
    , p_height pls_integer := null
    );
  --
  function read
    ( p_xlsx           blob
    , p_sheets         varchar2 := null
    , p_cell           varchar2 := null
    , p_include_clobs  varchar2 := null
    , p_add_empty_cols varchar2 := null
    )
  return tp_all_cells pipelined;
  --
  function get_sheet_names( p_xlsx blob )
  return sheet_names;
  --
  function file2blob( p_dir varchar2, p_file_name varchar2 )
  return blob;
  --
  function get_version
  return varchar2;
  --
end as_xlsx;
/
