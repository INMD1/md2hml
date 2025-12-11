
# -*- coding: utf-8 -*-

import sys
import re
import io

import os
import base64
import xml.sax.saxutils as saxutils
from datetime import datetime
from pytz import timezone

datetime.now(timezone('Asia/Seoul'))

TXT_END='</CHAR></TEXT></P>'

H1_START='<P ParaShape="20" Style="0"><TEXT CharShape="10"><CHAR>'
H2_START='<P ParaShape="20" Style="0"><TEXT CharShape="11"><CHAR>'
H3_START='<P ParaShape="20" Style="0"><TEXT CharShape="12"><CHAR>'
H4_START='<P ParaShape="20" Style="0"><TEXT CharShape="13"><CHAR>'
H5_START='<P ParaShape="20" Style="0"><TEXT CharShape="14"><CHAR>'
H6_START='<P ParaShape="20" Style="0"><TEXT CharShape="7"><CHAR>'

BIN_DATA_ENTRIES = []

def process_inline(text):
    # Bold: **text**
    # Replacement breaks current CHAR/TEXT block, inserts Bold TEXT block, then resumes Normal TEXT block
    # Note: CharShape="4" is Bold, CharShape="0" is Normal
    # Regex ensures content doesn't start/end with space to avoid matching list bullets or disjointed *
    text = re.sub(r'\*\*(?=\S)(.+?)(?<=\S)\*\*', r'</CHAR></TEXT><TEXT CharShape="4"><CHAR>\1</CHAR></TEXT><TEXT CharShape="0"><CHAR>', text)
    
    # Italic: *text*
    # CharShape="6" is Italic (newly added)
    # Lookbehind (?<!\*) and lookahead (?!\*) to avoid matching ** (bold) parts as italic
    text = re.sub(r'(?<!\*)\*(?=\S)(.+?)(?<=\S)\*(?!\*)', r'</CHAR></TEXT><TEXT CharShape="6"><CHAR>\1</CHAR></TEXT><TEXT CharShape="0"><CHAR>', text)

    # Bold: __text__ (Underscore)
    text = re.sub(r'__(?=\S)(.+?)(?<=\S)__', r'</CHAR></TEXT><TEXT CharShape="4"><CHAR>\1</CHAR></TEXT><TEXT CharShape="0"><CHAR>', text)
    
    # Italic: _text_ (Underscore)
    # Avoid matching inside words like some_variable_name if handled strictly, but standard simplified:
    text = re.sub(r'(?<!_)_(?=\S)(.+?)(?<=\S)_(?!_)', r'</CHAR></TEXT><TEXT CharShape="6"><CHAR>\1</CHAR></TEXT><TEXT CharShape="0"><CHAR>', text)

    # Inline Code: `text` -> CharShape="16" (Monospace)
    text = re.sub(r'`(.*?)`', r'</CHAR></TEXT><TEXT CharShape="16"><CHAR>\1</CHAR></TEXT><TEXT CharShape="0"><CHAR>', text)

    return text




def process_code_block(text):
    # Native HWP Table 1x1 for Code Block
    # text contains the content. Split by lines.
    lines = text.strip().split('\n')
    
    # We construct the content inside the CELL
    # Use ParaShape="0" for simple text inside (or a custom one)
    # CharShape="7" or "16" (Monospace)
    cell_content = ""
    for line in lines:
        cell_content += f'<P ParaShape="0" Style="0"><TEXT CharShape="16"><CHAR>{saxutils.escape(line)}</CHAR></TEXT></P>'
    
    # Wrap in TABLE XML structure
    # Using generic Size/Position values from reference
    # Must wrap in <P><TEXT>...</TEXT></P> for HWPML validity
    res = f"""<P ParaShape="0" Style="0"><TEXT CharShape="0"><TABLE BorderFill="3" CellSpacing="0" ColCount="1" PageBreak="Cell" RepeatHeader="true" RowCount="1">
<SHAPEOBJECT Lock="false" NumberingType="Table" TextWrap="TopAndBottom" ZOrder="1">
<SIZE Height="0" HeightRelTo="Absolute" Protect="false" Width="41954" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="false" AllowOverlap="false" FlowWithText="true" HoldAnchorAndSO="false" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="false" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="283" Left="283" Right="283" Top="283"/>
</SHAPEOBJECT>
<INSIDEMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<ROW>
<CELL BorderFill="3" ColAddr="0" ColSpan="1" Dirty="false" Editable="false" HasMargin="false" Header="false" Height="282" Protect="false" RowAddr="0" RowSpan="1" Width="41954">
<CELLMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<PARALIST LineWrap="Break" LinkListID="0" LinkListIDNext="0" TextDirection="0" VertAlign="Center">
{cell_content}
</PARALIST>
</CELL>
</ROW>
</TABLE><CHAR/></TEXT></P>"""
    return res

def process_table_native(table_block):
    # table_block is the raw markdown table string
    lines = table_block.strip().split('\n')
    if len(lines) < 2: return table_block # Not a valid table

    # Parse Header
    # Remove leading/trailing | and split
    header_cells = [c.strip() for c in lines[0].strip('|').split('|')]
    col_count = len(header_cells)
    
    # Skip separator line (lines[1])
    
    # Parse Rows
    rows_data = []
    # Add Header as first row data (HWP renders it same, styling optional)
    rows_data.append(header_cells)
    
    for i in range(2, len(lines)):
        row_cells = [c.strip() for c in lines[i].strip('|').split('|')]
        # Handle cell count mismatch
        if len(row_cells) < col_count:
            row_cells += [''] * (col_count - len(row_cells))
        rows_data.append(row_cells[:col_count])

    row_count = len(rows_data)
    
    # Build XML
    # Standard width for A4 roughly 42000 units total. Split evenly.
    total_width = 41954
    col_width = int(total_width / col_count)
    
    hml_rows = ""
    for r_idx, row in enumerate(rows_data):
        cells_xml = ""
        for c_idx, cell_text in enumerate(row):
            # Process inline formatting in cell text
            # We can lightly recurse or just text escape
            # Applying simpler logic: escape, then generic wrapper
            # Use CharShape="0" for normal text
            c_txt = saxutils.escape(cell_text)
            
            # Simple bold for header (first row) if desired, but sticking to basic first
            
            cells_xml += f"""<CELL BorderFill="3" ColAddr="{c_idx}" ColSpan="1" Dirty="false" Editable="false" HasMargin="false" Header="false" Height="282" Protect="false" RowAddr="{r_idx}" RowSpan="1" Width="{col_width}">
<CELLMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<PARALIST LineWrap="Break" LinkListID="0" LinkListIDNext="0" TextDirection="0" VertAlign="Center">
<P ParaShape="0" Style="0"><TEXT CharShape="0"><CHAR>{c_txt}</CHAR></TEXT></P>
</PARALIST>
</CELL>"""
        hml_rows += f"<ROW>{cells_xml}</ROW>"

    res = f"""<P ParaShape="0" Style="0"><TEXT CharShape="0"><TABLE BorderFill="3" CellSpacing="0" ColCount="{col_count}" PageBreak="Cell" RepeatHeader="true" RowCount="{row_count}">
<SHAPEOBJECT Lock="false" NumberingType="Table" TextWrap="TopAndBottom" ZOrder="0">
<SIZE Height="0" HeightRelTo="Absolute" Protect="false" Width="{total_width}" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="false" AllowOverlap="false" FlowWithText="true" HoldAnchorAndSO="false" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="false" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="283" Left="283" Right="283" Top="283"/>
</SHAPEOBJECT>
<INSIDEMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
{hml_rows}
</TABLE><CHAR/></TEXT></P>"""
    return res

def process_image(alt, path):
    # check if file exists
    if not os.path.exists(path):
        return f'[Image not found: {path}]'
    
    # Read and encode
    try:
        with open(path, "rb") as image_file:
             encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
    except Exception as e:
        return f'[Image error: {e}]'
    
    # Add to global BINDATA list
    global BIN_DATA_ENTRIES
    bin_id = len(BIN_DATA_ENTRIES) + 1
    ext = os.path.splitext(path)[1].replace('.', '').lower()
    if ext == 'jpeg': ext = 'jpg'
    
    BIN_DATA_ENTRIES.append(f'<BINITEM Id="{bin_id}" BinData="{encoded_string}" Format="{ext}" Type="Embedding" />')
    
    # Return Placeholder (until exact PICTURE tag syntax is confirmed)
    # Using a recognizable placeholder
    return f'[Image Embedded: {alt} (ID: {bin_id})]'


def get_header(bindata_list=""):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<HWPML Style="embed" SubVersion="10.0.0.0" Version="2.91">
    <HEAD SecCnt="1">
        <DOCSUMMARY>
            <TITLE>Report_</TITLE>
            <AUTHOR>___AUTHOR___</AUTHOR>
            <DATE>{datetime.now(timezone('Asia/Seoul')).strftime('%Y년 %m월 %d일 %A 오전 %I:%M:%S')}</DATE>
        </DOCSUMMARY>
        <DOCSETTING>
            <BEGINNUMBER Endnote="1" Equation="1" Footnote="1" Page="1" Picture="1" Table="1" />
            <CARETPOS List="0" Para="36" Pos="0" />
        </DOCSETTING>
        <MAPPINGTABLE>
            <FACENAMELIST>
                <FONTFACE Count="2" Lang="Hangul">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="Latin">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="Hanja">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="Japanese">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="Other">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="Symbol">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="User">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
            </FACENAMELIST>
            <BORDERFILLLIST Count="3">
                <BORDERFILL BackSlash="0" BreakCellSeparateLine="0" CenterLine="0"
                    CounterBackSlash="0" CounterSlash="0" CrookedSlash="0" Id="1" Shadow="false"
                    Slash="0" ThreeD="false">
                    <LEFTBORDER Type="None" Width="0.1mm" />
                    <RIGHTBORDER Type="None" Width="0.1mm" />
                    <TOPBORDER Type="None" Width="0.1mm" />
                    <BOTTOMBORDER Type="None" Width="0.1mm" />
                    <DIAGONAL Type="Solid" Width="0.1mm" />
                </BORDERFILL>
                <BORDERFILL BackSlash="0" BreakCellSeparateLine="0" CenterLine="0"
                    CounterBackSlash="0" CounterSlash="0" CrookedSlash="0" Id="2" Shadow="false"
                    Slash="0" ThreeD="false">
                    <LEFTBORDER Type="None" Width="0.1mm" />
                    <RIGHTBORDER Type="None" Width="0.1mm" />
                    <TOPBORDER Type="None" Width="0.1mm" />
                    <BOTTOMBORDER Type="None" Width="0.1mm" />
                    <DIAGONAL Type="Solid" Width="0.1mm" />
                    <FILLBRUSH>
                        <WINDOWBRUSH Alpha="0" FaceColor="4294967295" HatchColor="10066329" />
                    </FILLBRUSH>
                </BORDERFILL>
                <BORDERFILL BackSlash="0" BreakCellSeparateLine="0" CenterLine="0"
                    CounterBackSlash="0" CounterSlash="0" CrookedSlash="0" Id="3" Shadow="false"
                    Slash="0" ThreeD="false">
                    <LEFTBORDER Type="Solid" Width="0.12mm" />
                    <RIGHTBORDER Type="Solid" Width="0.12mm" />
                    <TOPBORDER Type="Solid" Width="0.12mm" />
                    <BOTTOMBORDER Type="Solid" Width="0.12mm" />
                    <DIAGONAL Type="Solid" Width="0.1mm" />
                </BORDERFILL>
            </BORDERFILLLIST>
            <CHARSHAPELIST Count="17">
                <CHARSHAPE BorderFillId="2" Height="1000" Id="0" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="1" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="900" Id="2" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="900" Id="3" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="900" Id="4" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="-5" Hanja="-5" Japanese="-5" Latin="-5" Other="-5"
                        Symbol="-5" User="-5" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1600" Id="5" ShadeColor="4294967295" SymMark="0"
                    TextColor="11891758" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1100" Id="6" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="7" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1900" Id="8" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="2400" Id="9" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="3200" Id="10" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="2400" Id="11" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1900" Id="12" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1600" Id="13" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1200" Id="14" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="15" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <ITALIC />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="16" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                    <UNDERLINE Color="0" Shape="Solid" Type="Bottom" />
                </CHARSHAPE>
            </CHARSHAPELIST>
            <TABDEFLIST Count="3">
                <TABDEF AutoTabLeft="false" AutoTabRight="false" Id="0" />
                <TABDEF AutoTabLeft="true" AutoTabRight="false" Id="1" />
                <TABDEF AutoTabLeft="false" AutoTabRight="true" Id="2" />
            </TABDEFLIST>
            <NUMBERINGLIST Count="2">
                <NUMBERING Id="1" Start="0">
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="1" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="2"
                        NumFormat="HangulSyllable" Start="1" TextOffset="50"
                        TextOffsetType="percent" UseInstWidth="true" WidthAdjust="0">^2.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="3" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^3)</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="4"
                        NumFormat="HangulSyllable" Start="1" TextOffset="50"
                        TextOffsetType="percent" UseInstWidth="true" WidthAdjust="0">^4)</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="5" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">(^5)</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="6"
                        NumFormat="HangulSyllable" Start="1" TextOffset="50"
                        TextOffsetType="percent" UseInstWidth="true" WidthAdjust="0">(^6)</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="7" NumFormat="CircledDigit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^7</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="8"
                        NumFormat="CircledHangulSyllable" Start="1" TextOffset="50"
                        TextOffsetType="percent" UseInstWidth="true" WidthAdjust="0">^8</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="9" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0"></PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="10" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0"></PARAHEAD>
                </NUMBERING>
                <NUMBERING Id="2" Start="0">
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="1" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="2" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="3" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="4" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="5" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="6" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.^6.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="7" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.^6.^7.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="8" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.^6.^7.^8.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="9" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.^6.^7.^8.^9.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="10" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.^6.^7.^8.^9.^:.</PARAHEAD>
                </NUMBERING>
            </NUMBERINGLIST>
            <BULLETLIST Count="1">
                <BULLET Char="" Id="1">
                    <PARAHEAD Alignment="Left" AutoIndent="true" TextOffset="50"
                        TextOffsetType="percent" UseInstWidth="false" WidthAdjust="0" />
                </BULLET>
            </BULLETLIST>
            <PARASHAPELIST Count="26">
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="0" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="1" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="3000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="2" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="2000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="3" KeepLines="false"
                    KeepWithNext="false" Level="1" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="4000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="4" KeepLines="false"
                    KeepWithNext="false" Level="2" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="6000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="5" KeepLines="false"
                    KeepWithNext="false" Level="3" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="8000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="6" KeepLines="false"
                    KeepWithNext="false" Level="4" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="10000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="7" KeepLines="false"
                    KeepWithNext="false" Level="5" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="12000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="8" KeepLines="false"
                    KeepWithNext="false" Level="6" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="14000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="9" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="150" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="10" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="-2620" Left="0" LineSpacing="130" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Left" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="11" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="130" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Left" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="20"
                    FontLineHeight="false" HeadingType="None" Id="12" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="600" Prev="2400" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Left" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="13" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="2" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="1400" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Left" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="14" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="2" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="2200" LineSpacing="160" LineSpacingType="Percent"
                        Next="1400" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Left" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="15" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="2" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="4400" LineSpacing="160" LineSpacingType="Percent"
                        Next="1400" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="Outline" Id="16" KeepLines="false"
                    KeepWithNext="false" Level="8" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="18000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="Outline" Id="17" KeepLines="false"
                    KeepWithNext="false" Level="9" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="20000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="Outline" Id="18" KeepLines="false"
                    KeepWithNext="false" Level="7" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="16000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="19" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="150" LineSpacingType="Percent"
                        Next="1600" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="20" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" Heading="1" HeadingType="Bullet" Id="21"
                    KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break"
                    PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline"
                    WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" Heading="1" HeadingType="Bullet" Id="22"
                    KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break"
                    PageBreakBefore="false" SnapToGrid="true" TabDef="1" VerAlign="Baseline"
                    WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="2000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="23" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="None" Id="24" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="2000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" Heading="1" HeadingType="Number" Id="25"
                    KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break"
                    PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline"
                    WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
            </PARASHAPELIST>
            <STYLELIST Count="22">
                <STYLE CharShape="0" EngName="Normal" Id="0" LangId="1042" LockForm="0" Name="바탕글"
                    NextStyle="0" ParaShape="0" Type="Para" />
                <STYLE CharShape="0" EngName="Body" Id="1" LangId="1042" LockForm="0" Name="본문"
                    NextStyle="1" ParaShape="1" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 1" Id="2" LangId="1042" LockForm="0"
                    Name="개요 1" NextStyle="2" ParaShape="2" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 2" Id="3" LangId="1042" LockForm="0"
                    Name="개요 2" NextStyle="3" ParaShape="3" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 3" Id="4" LangId="1042" LockForm="0"
                    Name="개요 3" NextStyle="4" ParaShape="4" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 4" Id="5" LangId="1042" LockForm="0"
                    Name="개요 4" NextStyle="5" ParaShape="5" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 5" Id="6" LangId="1042" LockForm="0"
                    Name="개요 5" NextStyle="6" ParaShape="6" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 6" Id="7" LangId="1042" LockForm="0"
                    Name="개요 6" NextStyle="7" ParaShape="7" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 7" Id="8" LangId="1042" LockForm="0"
                    Name="개요 7" NextStyle="8" ParaShape="8" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 8" Id="9" LangId="1042" LockForm="0"
                    Name="개요 8" NextStyle="9" ParaShape="18" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 9" Id="10" LangId="1042" LockForm="0"
                    Name="개요 9" NextStyle="10" ParaShape="16" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 10" Id="11" LangId="1042" LockForm="0"
                    Name="개요 10" NextStyle="11" ParaShape="17" Type="Para" />
                <STYLE CharShape="1" EngName="Page Number" Id="12" LangId="1042" LockForm="0"
                    Name="쪽 번호" NextStyle="0" Type="Char" />
                <STYLE CharShape="2" EngName="Header" Id="13" LangId="1042" LockForm="0" Name="머리말"
                    NextStyle="13" ParaShape="9" Type="Para" />
                <STYLE CharShape="3" EngName="Footnote" Id="14" LangId="1042" LockForm="0" Name="각주"
                    NextStyle="14" ParaShape="10" Type="Para" />
                <STYLE CharShape="3" EngName="Endnote" Id="15" LangId="1042" LockForm="0" Name="미주"
                    NextStyle="15" ParaShape="10" Type="Para" />
                <STYLE CharShape="4" EngName="Memo" Id="16" LangId="1042" LockForm="0" Name="메모"
                    NextStyle="16" ParaShape="11" Type="Para" />
                <STYLE CharShape="5" EngName="TOC Heading" Id="17" LangId="1042" LockForm="0"
                    Name="차례 제목" NextStyle="17" ParaShape="12" Type="Para" />
                <STYLE CharShape="6" EngName="TOC 1" Id="18" LangId="1042" LockForm="0" Name="차례 1"
                    NextStyle="18" ParaShape="13" Type="Para" />
                <STYLE CharShape="6" EngName="TOC 2" Id="19" LangId="1042" LockForm="0" Name="차례 2"
                    NextStyle="19" ParaShape="14" Type="Para" />
                <STYLE CharShape="6" EngName="TOC 3" Id="20" LangId="1042" LockForm="0" Name="차례 3"
                    NextStyle="20" ParaShape="15" Type="Para" />
                <STYLE CharShape="0" EngName="Caption" Id="21" LangId="1042" LockForm="0" Name="캡션"
                    NextStyle="21" ParaShape="19" Type="Para" />
            </STYLELIST>
        </MAPPINGTABLE>
        <COMPATIBLEDOCUMENT TargetProgram="None">
            <LAYOUTCOMPATIBILITY AdjustBaselineInFixedLinespacing="false"
                AdjustBaselineOfObjectToBottom="false" AdjustLineheightToFont="false"
                AdjustMarginFromAdjustLineheight="false" AdjustParaBorderOffsetWithBorder="false"
                AdjustParaBorderfillToSpacing="false" AdjustVertPosOfLine="false"
                ApplyAtLeastToPercent100Pct="false" ApplyCharSpacingToCharGrid="false"
                ApplyExtendHeaderFooterEachSection="false" ApplyFontWeightToBold="false"
                ApplyFontspaceToLatin="false" ApplyMinColumnWidthTo1mm="false"
                ApplyNextspacingOfLastPara="false" ApplyParaBorderToOutside="false"
                ApplyPrevspacingBeneathObject="false" ApplyTabPosBasedOnSegment="false"
                BaseCharUnitOfIndentOnFirstChar="false" BaseCharUnitOnEAsian="false"
                BaseLinespacingOnLinegrid="false" BreakTabOverLine="false"
                ConnectParaBorderfillOfEqualBorder="false" DoNotAdjustEmptyAnchorLine="false"
                DoNotAdjustWordInJustify="false" DoNotAlignLastForbidden="false"
                DoNotAlignLastPeriod="false" DoNotAlignWhitespaceOnRight="false"
                DoNotApplyAutoSpaceEAsianEng="false" DoNotApplyAutoSpaceEAsianNum="false"
                DoNotApplyColSeparatorAtNoGap="false" DoNotApplyExtensionCharCompose="false"
                DoNotApplyGridInHeaderFooter="false" DoNotApplyHeaderFooterAtNoSpace="false"
                DoNotApplyImageEffect="false" DoNotApplyLinegridAtNoLinespacing="false"
                DoNotApplyShapeComment="false" DoNotApplyStrikeoutWithUnderline="false"
                DoNotApplyVertOffsetOfForward="false" DoNotApplyWhiteSpaceHeight="false"
                DoNotFormattingAtBeneathAnchor="false" DoNotHoldAnchorOfTable="false"
                ExtendLineheightToOffset="false" ExtendLineheightToParaBorderOffset="false"
                ExtendVertLimitToPageMargins="false" FixedUnderlineWidth="false"
                OverlapBothAllowOverlap="false" TreatQuotationAsLatin="false"
                UseInnerUnderline="false" UseLowercaseStrikeout="false" />
        </COMPATIBLEDOCUMENT>
    </HEAD>
    <BODY>
        <SECTION Id="0">"""

FOOTER="""</SECTION></BODY><TAIL></TAIL></HWPML>"""

#인트로는 없습(원래 md파일 유지)
INTRO="""
"""

fIn  = "README.md"
fOut = "readme.hml"

if len(sys.argv) > 1:
    fIn = sys.argv[1]

if len(sys.argv) > 2:
    fOut = sys.argv[2]

print("in:%s -> out:%s"%(fIn, fOut))


with open(fIn, 'r') as content_file:
    ret = content_file.read()

ret = re.sub("\r", "", ret)


# ---------------------------------------------------------
# Processing Logic
# ---------------------------------------------------------

# Ensure content starts with newline for regex matching
ret = "\n" + ret

# 1. XML Escape Text Content
ret = saxutils.escape(ret)

# Parse metadata after escaping (so we match escaped keys if necessary, though unlikely for standard markdown)
t_match = re.search("(?<=title: ).*", ret)
t = t_match.group(0) if t_match else "No Title"

a_match = re.search("(?<=author: ).*", ret)
a = a_match.group(0) if a_match else "Unknown"

d_match = re.search("(?<=date: ).*", ret)
d = d_match.group(0) if d_match else "Unknown"

# Remove Front Matter
ret = re.sub("---(.|\n)*?---", "", ret, count=1)


# 2. Images: ![alt](path)
# Must be done before text wrapping
# Return text placeholder to avoid nested <P> tags invalidating XML
ret = re.sub(r'!\[(.*?)\]\((.*?)\)', lambda m: process_image(m.group(1), m.group(2)), ret)


# 3. Code Blocks (Pre-processing)
# Must be done before text wrapping to apply special ParaShape
# Ensure it matches start of line to avoid inline issues
# Update: Now returns a TABLE XML string.
# We need to ensure this XML string is NOT wrapped in <P> by the Text Wrapping step (Step 7)
# We can wrap it in a placeholder, or rely on Text Wrapping respecting tags.
# Step 7 wraps lines not starting with <. Our Table string starts with <TABLE, so it should be safe.
ret = re.sub(r'(?m)^```(.*?)```', lambda m: process_code_block(m.group(1)), ret, flags=re.DOTALL)


# 4. Page Break / HR: ---
# Map to ParaShape with Bottom Border (ParaShape="10")
ret = re.sub(r'\n---', r'\n<P ParaShape="10" Style="0"><TEXT CharShape="0"><CHAR></CHAR></TEXT></P>', ret)


# 5. Links: [text](url) -> text (url)
ret = re.sub(r'\[(.*?)\]\((.*?)\)', r'\1 (\2)', ret)

# 6. General Markdown (Lists, Headers)
BASE_INST_ID = 3102933706

def replace_md_list(m):
    global BASE_INST_ID
    spaces = m.group(1)
    bullet = m.group(2)

    depth = len(spaces) // 4         # 4 spaces = 1 depth
    para = 21 + depth                # ParaShape 21~26

    if para == 22:
        BASE_INST_ID = BASE_INST_ID + 1
        inst_attr = f' InstId="{BASE_INST_ID}"'
    else:
        inst_attr = ""

    return f'\n<P{inst_attr} ParaShape="{para}" Style="0"><TEXT CharShape="0"><CHAR>{bullet} '

ret = re.sub(
    r'\n( {0,24})([-+*]) ',
    lambda m: replace_md_list(m).replace("* ", ""),
    ret
)

# Ordered Lists (1. item) -> Use ParaShape 25 (Indented)
ret = re.sub(r'\n(\d+)\. ', r'\n<P ParaShape="25" Style="0"><TEXT CharShape="0"><CHAR>', ret)

# Nested Ordered Lists (    1. item) -> Use InstId
ret = re.sub(r'\n    (\d+)\. ', r'\n<P InstId="95545006\1" ParaShape="2" Style="2"><TEXT CharShape="0"><CHAR>', ret)

# Task Lists
ret = re.sub(r'\[ \]', r'☐', ret)
ret = re.sub(r'\[x\]', r'☑', ret)
ret = re.sub(r'\[X\]', r'☑', ret)

# Blockquotes (> text) -> Use ParaShape 11 (Indented) + CharShape 8 (Gray/Italic)
# Handle > and &gt; (due to escaping)
ret = re.sub(r'\n(&gt;|>) (.*)', r'\n<P ParaShape="11" Style="0"><TEXT CharShape="8"><CHAR>\2</CHAR></TEXT></P>', ret)

# Tables (Native XML)
# Detect full table blocks. Regex looks for lines starting with | and ending with | (multiline)
# We capture the whole block and pass to process_table_native
# Note: This is complex with regex. We assume a table block is a contiguous set of lines starting with |
ret = re.sub(r'(?m)((?:^\|.*\|$\n?)+)', lambda m: process_table_native(m.group(1)), ret)

ret = re.sub("\n# "     , "\n"+H1_START, ret)
ret = re.sub("\n## "    , "\n"+H2_START, ret)
ret = re.sub("\n### "   , "\n"+H3_START, ret)
ret = re.sub("\n#### "  , "\n"+H4_START, ret)
ret = re.sub("\n##### " , "\n"+H5_START, ret)
ret = re.sub("\n###### ", "\n"+H6_START, ret)


# 7. Text Wrapping (The Catch-All)
# Wrap any line that doesn't start with an HML Tag (<)
ret = re.sub("\n(?=[^<])", "\n", ret)

# Close Tags
# Use negative lookbehind (?<!</P>) to avoid adding TXT_END to lines that are already closed (like Code Blocks or HRs)
# Also prevent adding it to lines ending with > (XML structure like TABLE, ROW, etc)
ret = re.sub(
    r'(<CHAR>)([^<]*?)(?=\s*(?:</TEXT>|</P>|<P|$))',
    r'\1\2</CHAR>',
    ret,
    flags=re.DOTALL
)

# 그 다음 </TEXT></P> 추가
ret = re.sub(
    r'(</CHAR>)(?!\s*</TEXT>)',
    r'\1</TEXT></P>',
    ret
)

# 파일 끝에서도 닫히지 않은 경우 처리
ret = re.sub(
    r'(<CHAR>[^<]*)$',
    r'\1</CHAR></TEXT></P>',
    ret,
    flags=re.MULTILINE
)
# ret = re.sub(r"(?<!</P>)(?<!>)\n", TXT_END+"\n", ret)
# ret = re.sub("^"+TXT_END, "", ret)



# Process Inline Formatting (Bold, Italic)
# Only apply to CharShape="0" (Normal Text) to avoid messing up Headers
# Call process_inline using re.sub with callback
ret = re.sub(r'(<TEXT CharShape="0"><CHAR>)(.*?)(</CHAR>)', lambda m: m.group(1) + process_inline(m.group(2)) + m.group(3), ret, flags=re.DOTALL)
# Call process_inline using re.sub with callback
ret = re.sub(r'(<TEXT CharShape="0"><CHAR>)(.*?)(</CHAR>)', lambda m: m.group(1) + process_inline(m.group(2)) + m.group(3), ret, flags=re.DOTALL)

# Use dynamic header with BinData
bindata_str = '<BINDATALIST Count="' + str(len(BIN_DATA_ENTRIES)) + '">' + "".join(BIN_DATA_ENTRIES) + '</BINDATALIST>' if BIN_DATA_ENTRIES else ""
header_content = get_header(bindata_str)

ret = header_content + INTRO + ret + FOOTER

ret = re.sub("___TITLE___" , t, ret)
ret = re.sub("___AUTHOR___", a, ret)
ret = re.sub("___DATE___"  , d, ret)

with io.open(fOut, "w", encoding="utf-8") as f:
    f.write(ret)
    f.close()

