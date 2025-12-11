
# -*- coding: utf-8 -*-

import sys
import re
import io

import os
import base64
import xml.sax.saxutils as saxutils

TXT_START='<P ParaShape="0" Style="0"><TEXT CharShape="0"><CHAR>'
TXT_END='</CHAR></TEXT></P>'

L1_START='<P ParaShape="1" Style="0"><TEXT CharShape="0"><CHAR>* '
L2_START='<P ParaShape="2" Style="0"><TEXT CharShape="0"><CHAR>  - '
L3_START='<P ParaShape="3" Style="0"><TEXT CharShape="0"><CHAR>    + '
L4_START='<P ParaShape="4" Style="0"><TEXT CharShape="0"><CHAR>      * '
L5_START='<P ParaShape="5" Style="0"><TEXT CharShape="0"><CHAR>        - '
L6_START='<P ParaShape="6" Style="0"><TEXT CharShape="0"><CHAR>          + '
H1_START='<P ParaShape="0" Style="0"><TEXT CharShape="1"><CHAR>'
H2_START='<P ParaShape="0" Style="0"><TEXT CharShape="2"><CHAR>'
H3_START='<P ParaShape="0" Style="0"><TEXT CharShape="3"><CHAR>'
H4_START='<P ParaShape="0" Style="0"><TEXT CharShape="4"><CHAR>'

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
    return text



def process_code_block(text):
    # ```code```
    # We match the whole block.
    # We want to replace it with ParaShape="9" (Code Block Style)
    # Content must be split by newline and each line wrapped in <P>...</P>
    # Note: text contains the content inside ``` ```
    lines = text.split('\n')
    res = ""
    for line in lines:
        if not line: continue
        # CharShape="7" is Monospace
        # ParaShape="9" is Indented + Background (if we can)
        res += f'<P ParaShape="9" Style="0"><TEXT CharShape="7"><CHAR>{line}</CHAR></TEXT></P>\n'
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
    return """<?xml version="1.0" encoding="UTF-8" standalone="no"?><HWPML Style="embed" SubVersion="9.0.1.0" Version="2.9"><HEAD SecCnt="1">
<DOCSUMMARY><TITLE>___TITLE___</TITLE><AUTHOR>___AUTHOR___</AUTHOR><DATE>___DATE___</DATE></DOCSUMMARY><DOCSETTING>
<BEGINNUMBER Endnote="1" Equation="1" Footnote="1" Page="1" Picture="1" Table="1" /><CARETPOS List="0" Para="20" Pos="6" /></DOCSETTING><MAPPINGTABLE>""" + bindata_list + """<FACENAMELIST><FONTFACE Count="2" Lang="Hangul"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="Latin"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="Hanja"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="Japanese"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="Other"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="Symbol"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="User"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE></FACENAMELIST>
<BORDERFILLLIST Count="5">
    <BORDERFILL Id="1" ThreeD="false" Shadow="false" CenterLine="None" BackgroundBrushId="0" />
    <BORDERFILL Id="2" ThreeD="false" Shadow="false" CenterLine="None" BackgroundBrushId="0" />
    <BORDERFILL Id="3" ThreeD="false" Shadow="false" CenterLine="None" BackgroundBrushId="0">
        <FILLBRUSH><WINDOWBRUSH FaceColor="14737632" HatchColor="0" HatchStyle="None" Alpha="0" /></FILLBRUSH>
    </BORDERFILL>
    <BORDERFILL Id="4" ThreeD="false" Shadow="false" CenterLine="None" BackgroundBrushId="0">
        <BOTTOMBORDER Type="Solid" Width="0.12mm" Color="0" />
    </BORDERFILL>
</BORDERFILLLIST>
<CHARSHAPELIST Count="9">
    <CHARSHAPE BorderFillId="2" Height="1000" Id="0" ShadeColor="4294967295" SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
        <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1" User="1" />
        <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
    </CHARSHAPE>
    <CHARSHAPE BorderFillId="2" Height="2000" Id="1" ShadeColor="4294967295" SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
        <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1" User="1" />
        <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <BOLD />
    </CHARSHAPE>
    <CHARSHAPE BorderFillId="2" Height="1600" Id="2" ShadeColor="4294967295" SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
        <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1" User="1" />
        <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <BOLD />
    </CHARSHAPE>
    <CHARSHAPE BorderFillId="2" Height="1200" Id="3" ShadeColor="4294967295" SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
        <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1" User="1" />
        <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <BOLD />
    </CHARSHAPE>
    <CHARSHAPE BorderFillId="2" Height="1000" Id="4" ShadeColor="4294967295" SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
        <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1" User="1" />
        <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <BOLD />
    </CHARSHAPE>
    <CHARSHAPE BorderFillId="2" Height="2500" Id="5" ShadeColor="4294967295" SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
        <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1" User="1" />
        <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <BOLD />
    </CHARSHAPE>
    <CHARSHAPE BorderFillId="2" Height="1000" Id="6" ShadeColor="4294967295" SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
        <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1" User="1" />
        <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <ITALIC />
    </CHARSHAPE>
    <CHARSHAPE BorderFillId="2" Height="1000" Id="7" ShadeColor="4294967295" SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
        <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
    </CHARSHAPE>
    <CHARSHAPE BorderFillId="2" Height="1000" Id="8" ShadeColor="4294967295" SymMark="0" TextColor="8421504" UseFontSpace="false" UseKerning="false">
        <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1" User="1" />
        <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <ITALIC />
    </CHARSHAPE>
</CHARSHAPELIST>
<PARASHAPELIST Count="13">
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="0" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="1" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="-2104" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="2" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="-4174" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="3" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="-6244" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="4" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="-8314" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="5" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="-10384" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="6" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="-12454" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Center" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0" FontLineHeight="false" HeadingType="None" Id="7" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Right" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0" FontLineHeight="false" HeadingType="None" Id="8" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="9" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="2000" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="3" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="10" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="4" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="11" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="0" Left="2000" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="12" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="-2104" Left="2104" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
</PARASHAPELIST>

</MAPPINGTABLE><COMPATIBLEDOCUMENT TargetProgram="None"><LAYOUTCOMPATIBILITY AdjustBaselineInFixedLinespacing="false" AdjustBaselineOfObjectToBottom="false" AdjustLineheightToFont="false" AdjustMarginFromAdjustLineheight="false" AdjustParaBorderOffsetWithBorder="false" AdjustParaBorderfillToSpacing="false" AdjustVertPosOfLine="false" ApplyAtLeastToPercent100Pct="false" ApplyCharSpacingToCharGrid="false" ApplyExtendHeaderFooterEachSection="false" ApplyFontWeightToBold="false" ApplyFontspaceToLatin="false" ApplyMinColumnWidthTo1mm="false" ApplyNextspacingOfLastPara="false" ApplyParaBorderToOutside="false" ApplyPrevspacingBeneathObject="false" ApplyTabPosBasedOnSegment="false" BaseCharUnitOfIndentOnFirstChar="false" BaseCharUnitOnEAsian="false" BaseLinespacingOnLinegrid="false" BreakTabOverLine="false" ConnectParaBorderfillOfEqualBorder="false" DoNotAdjustEmptyAnchorLine="false" DoNotAdjustWordInJustify="false" DoNotAlignLastForbidden="false" DoNotAlignLastPeriod="false" DoNotAlignWhitespaceOnRight="false" DoNotApplyAutoSpaceEAsianEng="false" DoNotApplyAutoSpaceEAsianNum="false" DoNotApplyColSeparatorAtNoGap="false" DoNotApplyExtensionCharCompose="false" DoNotApplyGridInHeaderFooter="false" DoNotApplyHeaderFooterAtNoSpace="false" DoNotApplyImageEffect="false" DoNotApplyLinegridAtNoLinespacing="false" DoNotApplyShapeComment="false" DoNotApplyStrikeoutWithUnderline="false" DoNotApplyVertOffsetOfForward="false" DoNotApplyWhiteSpaceHeight="false" DoNotFormattingAtBeneathAnchor="false" DoNotHoldAnchorOfTable="false" ExtendLineheightToOffset="false" ExtendLineheightToParaBorderOffset="false" ExtendVertLimitToPageMargins="false" FixedUnderlineWidth="false" OverlapBothAllowOverlap="false" TreatQuotationAsLatin="false" UseInnerUnderline="false" UseLowercaseStrikeout="false" /></COMPATIBLEDOCUMENT></HEAD><BODY><SECDEF CharGrid="0" FirstBorder="false" FirstFill="false" LineGrid="0" OutlineShape="1" SpaceColumns="1134" TabStop="8000" TextDirection="0" TextVerticalWidthHead="0"><STARTNUMBER Equation="0" Figure="0" Page="0" PageStartsOn="Both" Table="0" /><HIDE Border="false" EmptyLine="false" Footer="false" Header="false" MasterPage="false" PageNumPos="false" /><PAGEDEF GutterType="LeftOnly" Height="84188" Landscape="0" Width="59528"><PAGEMARGIN Bottom="4252" Footer="4252" Gutter="0" Header="4252" Left="8504" Right="8504" Top="5668" /></PAGEDEF><FOOTNOTESHAPE><AUTONUMFORMAT SuffixChar=")" Superscript="false" Type="Digit" /><NOTELINE Length="5cm" Type="Solid" Width="0.12mm" /><NOTESPACING AboveLine="850" BelowLine="567" BetweenNotes="283" /><NOTENUMBERING NewNumber="1" Type="Continuous" /><NOTEPLACEMENT BeneathText="false" Place="EachColumn" /></FOOTNOTESHAPE><ENDNOTESHAPE><AUTONUMFORMAT SuffixChar=")" Superscript="false" Type="Digit" /><NOTELINE Length="14692344" Type="Solid" Width="0.12mm" /><NOTESPACING AboveLine="850" BelowLine="567" BetweenNotes="0" /><NOTENUMBERING NewNumber="1" Type="Continuous" /><NOTEPLACEMENT BeneathText="false" Place="EndOfDocument" /></ENDNOTESHAPE><PAGEBORDERFILL BorferFill="1" FillArea="Paper" FooterInside="false" HeaderInside="false" TextBorder="true" Type="Both"><PAGEOFFSET Bottom="1417" Left="1417" Right="1417" Top="1417" /></PAGEBORDERFILL><PAGEBORDERFILL BorferFill="1" FillArea="Paper" FooterInside="false" HeaderInside="false" TextBorder="true" Type="Even"><PAGEOFFSET Bottom="1417" Left="1417" Right="1417" Top="1417" /></PAGEBORDERFILL><PAGEBORDERFILL BorferFill="1" FillArea="Paper" FooterInside="false" HeaderInside="false" TextBorder="true" Type="Odd"><PAGEOFFSET Bottom="1417" Left="1417" Right="1417" Top="1417" /></PAGEBORDERFILL></SECDEF><SECTION Id="0">"""

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

t_match = re.search("(?<=title: ).*", ret)
t = t_match.group(0) if t_match else "No Title"

a_match = re.search("(?<=author: ).*", ret)
a = a_match.group(0) if a_match else "Unknown"

d_match = re.search("(?<=date: ).*", ret)
d = d_match.group(0) if d_match else "Unknown"

# Remove Front Matter
ret = re.sub("---(.|\n)*---", "", ret, count=1)


# 2. Images: ![alt](path)
# Must be done before text wrapping
# Return text placeholder to avoid nested <P> tags invalidating XML
ret = re.sub(r'!\[(.*?)\]\((.*?)\)', lambda m: process_image(m.group(1), m.group(2)), ret)


# 3. Code Blocks (Pre-processing)
# Must be done before text wrapping to apply special ParaShape
# Ensure it matches start of line to avoid inline issues
ret = re.sub(r'(?m)^```(.*?)```', lambda m: process_code_block(m.group(1)), ret, flags=re.DOTALL)


# 4. Page Break / HR: ---
# Map to ParaShape with Bottom Border (ParaShape="10")
ret = re.sub(r'\n---', r'\n<P ParaShape="10" Style="0"><TEXT CharShape="0"><CHAR></CHAR></TEXT></P>', ret)


# 5. Links: [text](url) -> text (url)
ret = re.sub(r'\[(.*?)\]\((.*?)\)', r'\1 (\2)', ret)


# 6. General Markdown (Lists, Headers)
ret = re.sub("\n[-+*] "                    , "\n"+L1_START, ret)
ret = re.sub("\n    [-+*] "                , "\n"+L2_START, ret)
ret = re.sub("\n        [-+*] "            , "\n"+L3_START, ret)
ret = re.sub("\n            [-+*] "        , "\n"+L4_START, ret)
ret = re.sub("\n                [-+*] "    , "\n"+L5_START, ret)
ret = re.sub("\n                    [-+*] ", "\n"+L6_START, ret)

# Ordered Lists (1. item) -> Use ParaShape 12 (Indented)
ret = re.sub(r'\n(\d+)\. ', r'\n<P ParaShape="12" Style="0"><TEXT CharShape="0"><CHAR>\1. ', ret)

# Task Lists
ret = re.sub(r'\[ \]', r'☐', ret)
ret = re.sub(r'\[x\]', r'☑', ret)
ret = re.sub(r'\[X\]', r'☑', ret)

# Blockquotes (> text) -> Use ParaShape 11 (Indented) + CharShape 8 (Gray/Italic)
ret = re.sub(r'\n> (.*)', r'\n<P ParaShape="11" Style="0"><TEXT CharShape="8"><CHAR>\1</CHAR></TEXT></P>', ret)

# Tables (Simple Visual Fallback)
# Detect lines starting and ending with | and treat them as monospaced text to preserve alignment.
# We skip the separator line |---| because it looks ugly in plain text usually, or we keep it. Use code block style.
# Using lookahead to check if line looks like a table row
ret = re.sub(r'(?m)^(\|.*\|$)', r'<P ParaShape="9" Style="0"><TEXT CharShape="7"><CHAR>\1</CHAR></TEXT></P>', ret)

ret = re.sub("\n# "     , "\n"+H1_START, ret)
ret = re.sub("\n## "    , "\n"+H2_START, ret)
ret = re.sub("\n### "   , "\n"+H3_START, ret)
ret = re.sub("\n#### "  , "\n"+H4_START, ret)
ret = re.sub("\n##### " , "\n"+H4_START, ret)
ret = re.sub("\n###### ", "\n"+H4_START, ret)


# 7. Text Wrapping (The Catch-All)
# Wrap any line that doesn't start with an HML Tag (<)
ret = re.sub("\n(?=[^<])", "\n"+TXT_START, ret)

# Close Tags
# Use negative lookbehind (?<!</P>) to avoid adding TXT_END to lines that are already closed (like Code Blocks or HRs)
ret = re.sub(r"(?<!</P>)\n", TXT_END+"\n", ret)
ret = re.sub("^"+TXT_END, "", ret)



# Process Inline Formatting (Bold, Italic)
# Only apply to CharShape="0" (Normal Text) to avoid messing up Headers
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

