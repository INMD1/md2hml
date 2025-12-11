
import * as fs from 'fs';
import * as path from 'path';

// Constants
const TXT_START = '<P ParaShape="0" Style="0"><TEXT CharShape="0"><CHAR>';
const TXT_END = '</CHAR></TEXT></P>';

const L1_START = '<P ParaShape="1" Style="0"><TEXT CharShape="0"><CHAR>* ';
const L2_START = '<P ParaShape="2" Style="0"><TEXT CharShape="0"><CHAR>  - ';
const L3_START = '<P ParaShape="3" Style="0"><TEXT CharShape="0"><CHAR>    + ';
const L4_START = '<P ParaShape="4" Style="0"><TEXT CharShape="0"><CHAR>      * ';
const L5_START = '<P ParaShape="5" Style="0"><TEXT CharShape="0"><CHAR>        - ';
const L6_START = '<P ParaShape="6" Style="0"><TEXT CharShape="0"><CHAR>          + ';

const H1_START = '<P ParaShape="0" Style="0"><TEXT CharShape="1"><CHAR>';
const H2_START = '<P ParaShape="0" Style="0"><TEXT CharShape="2"><CHAR>';
const H3_START = '<P ParaShape="0" Style="0"><TEXT CharShape="3"><CHAR>';
const H4_START = '<P ParaShape="0" Style="0"><TEXT CharShape="4"><CHAR>';

const BIN_DATA_ENTRIES: string[] = [];

// Helper: XML Escape
function xmlEscape(str: string): string {
    if (!str) return "";
    return str.replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

// Helper: Process Image
function process_image(alt: string, imgPath: string): string {
    if (!fs.existsSync(imgPath)) {
        return `[Image not found: ${imgPath}]`;
    }

    try {
        const fileData = fs.readFileSync(imgPath);
        const encodedString = fileData.toString('base64');

        const bin_id = BIN_DATA_ENTRIES.length + 1;
        let ext = path.extname(imgPath).replace('.', '').toLowerCase();
        if (ext === 'jpeg') ext = 'jpg';

        BIN_DATA_ENTRIES.push(`<BINITEM Id="${bin_id}" BinData="${encodedString}" Format="${ext}" Type="Embedding" />`);

        return `[Image Embedded: ${alt} (ID: ${bin_id})]`;
    } catch (e: any) {
        return `[Image error: ${e.message}]`;
    }
}

// Helper: Process Code Block (1x1 Table)
function process_code_block(text: string): string {
    const lines = text.trim().split('\n');
    let cell_content = "";

    lines.forEach(line => {
        // Text is already escaped
        cell_content += `<P ParaShape="0" Style="0"><TEXT CharShape="16"><CHAR>${line}</CHAR></TEXT></P>`;
    });

    return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><TABLE BorderFill="3" CellSpacing="0" ColCount="1" PageBreak="Cell" RepeatHeader="true" RowCount="1">
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
${cell_content}
</PARALIST>
</CELL>
</ROW>
</TABLE><CHAR/></TEXT></P>`;
}

// Helper: Process Native Table
function process_table_native(table_block: string): string {
    const lines = table_block.trim().split('\n');
    if (lines.length < 2) return table_block;

    // Remove leading/trailing pipes
    const header_cells = lines[0].trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
    const col_count = header_cells.length;

    const rows_data: string[][] = [];
    rows_data.push(header_cells);

    for (let i = 2; i < lines.length; i++) {
        let row_cells = lines[i].trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
        if (row_cells.length < col_count) {
            row_cells = row_cells.concat(new Array(col_count - row_cells.length).fill(''));
        }
        rows_data.push(row_cells.slice(0, col_count));
    }

    const row_count = rows_data.length;
    const total_width = 41954;
    const col_width = Math.floor(total_width / col_count);

    let hml_rows = "";

    rows_data.forEach((row, r_idx) => {
        let cells_xml = "";
        row.forEach((cell_text, c_idx) => {
            // cell_text is already escaped
            cells_xml += `<CELL BorderFill="3" ColAddr="${c_idx}" ColSpan="1" Dirty="false" Editable="false" HasMargin="false" Header="false" Height="282" Protect="false" RowAddr="${r_idx}" RowSpan="1" Width="${col_width}">
<CELLMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<PARALIST LineWrap="Break" LinkListID="0" LinkListIDNext="0" TextDirection="0" VertAlign="Center">
<P ParaShape="0" Style="0"><TEXT CharShape="0"><CHAR>${cell_text}</CHAR></TEXT></P>
</PARALIST>
</CELL>`;
        });
        hml_rows += `<ROW>${cells_xml}</ROW>`;
    });

    return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><TABLE BorderFill="3" CellSpacing="0" ColCount="${col_count}" PageBreak="Cell" RepeatHeader="true" RowCount="${row_count}">
<SHAPEOBJECT Lock="false" NumberingType="Table" TextWrap="TopAndBottom" ZOrder="0">
<SIZE Height="0" HeightRelTo="Absolute" Protect="false" Width="${total_width}" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="false" AllowOverlap="false" FlowWithText="true" HoldAnchorAndSO="false" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="false" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="283" Left="283" Right="283" Top="283"/>
</SHAPEOBJECT>
<INSIDEMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
${hml_rows}
</TABLE><CHAR/></TEXT></P>`;
}

// Helper: Process Inline Formatting
function process_inline(text: string): string {
    // Bold: **text**
    text = text.replace(/\*\*(?=\S)(.+?)(?<=\S)\*\*/g, '</CHAR></TEXT><TEXT CharShape="4"><CHAR>$1</CHAR></TEXT><TEXT CharShape="0"><CHAR>');

    // Italic: *text*
    text = text.replace(/(?<!\*)\*(?=\S)(.+?)(?<=\S)\*(?!\*)/g, '</CHAR></TEXT><TEXT CharShape="6"><CHAR>$1</CHAR></TEXT><TEXT CharShape="0"><CHAR>');

    // Bold: __text__
    text = text.replace(/__(?=\S)(.+?)(?<=\S)__/g, '</CHAR></TEXT><TEXT CharShape="4"><CHAR>$1</CHAR></TEXT><TEXT CharShape="0"><CHAR>');

    // Italic: _text_
    text = text.replace(/(?<!_)_(?=\S)(.+?)(?<=\S)_(?!_)/g, '</CHAR></TEXT><TEXT CharShape="6"><CHAR>$1</CHAR></TEXT><TEXT CharShape="0"><CHAR>');

    // Inline Code: `text` -> CharShape="16"
    text = text.replace(/`(.*?)`/g, '</CHAR></TEXT><TEXT CharShape="16"><CHAR>$1</CHAR></TEXT><TEXT CharShape="0"><CHAR>');

    return text;
}

function get_header(bindata_str: string): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="no"?><HWPML Style="embed" SubVersion="9.0.1.0" Version="2.9"><HEAD SecCnt="1">
<DOCSUMMARY><TITLE>___TITLE___</TITLE><AUTHOR>___AUTHOR___</AUTHOR><DATE>___DATE___</DATE></DOCSUMMARY><DOCSETTING>
<BEGINNUMBER Endnote="1" Equation="1" Footnote="1" Page="1" Picture="1" Table="1" /><CARETPOS List="0" Para="20" Pos="6" /></DOCSETTING><MAPPINGTABLE>${bindata_str}<FACENAMELIST><FONTFACE Count="2" Lang="Hangul"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="Latin"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="Hanja"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="Japanese"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="Other"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="Symbol"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE><FONTFACE Count="2" Lang="User"><FONT Id="0" Name="함초롬돋움" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT><FONT Id="1" Name="함초롬바탕" Type="ttf"><TYPEINFO ArmStyle="1" Contrast="0" FamilyType="2" Letterform="1" Midline="1" Proportion="4" StrokeVariation="1" Weight="6" XHeight="1" /></FONT></FONTFACE></FACENAMELIST>
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
<CHARSHAPELIST Count="17">
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
    <CHARSHAPE BorderFillId="2" Height="1000" Id="16" ShadeColor="4294967295" SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
        <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1" User="1" />
        <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
        <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100" Symbol="100" User="100" />
        <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0" User="0" />
    </CHARSHAPE>
</CHARSHAPELIST>
<PARASHAPELIST Count="14">
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
    <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false" BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0" FontLineHeight="false" HeadingType="None" Id="13" KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
        <PARAMARGIN Indent="-2104" Left="4208" LineSpacing="160" LineSpacingType="Percent" Next="0" Prev="0" Right="0" />
        <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
    </PARASHAPE>
</PARASHAPELIST>

</MAPPINGTABLE><COMPATIBLEDOCUMENT TargetProgram="None"><LAYOUTCOMPATIBILITY AdjustBaselineInFixedLinespacing="false" AdjustBaselineOfObjectToBottom="false" AdjustLineheightToFont="false" AdjustMarginFromAdjustLineheight="false" AdjustParaBorderOffsetWithBorder="false" AdjustParaBorderfillToSpacing="false" AdjustVertPosOfLine="false" ApplyAtLeastToPercent100Pct="false" ApplyCharSpacingToCharGrid="false" ApplyExtendHeaderFooterEachSection="false" ApplyFontWeightToBold="false" ApplyFontspaceToLatin="false" ApplyMinColumnWidthTo1mm="false" ApplyNextspacingOfLastPara="false" ApplyParaBorderToOutside="false" ApplyPrevspacingBeneathObject="false" ApplyTabPosBasedOnSegment="false" BaseCharUnitOfIndentOnFirstChar="false" BaseCharUnitOnEAsian="false" BaseLinespacingOnLinegrid="false" BreakTabOverLine="false" ConnectParaBorderfillOfEqualBorder="false" DoNotAdjustEmptyAnchorLine="false" DoNotAdjustWordInJustify="false" DoNotAlignLastForbidden="false" DoNotAlignLastPeriod="false" DoNotAlignWhitespaceOnRight="false" DoNotApplyAutoSpaceEAsianEng="false" DoNotApplyAutoSpaceEAsianNum="false" DoNotApplyColSeparatorAtNoGap="false" DoNotApplyExtensionCharCompose="false" DoNotApplyGridInHeaderFooter="false" DoNotApplyHeaderFooterAtNoSpace="false" DoNotApplyImageEffect="false" DoNotApplyLinegridAtNoLinespacing="false" DoNotApplyShapeComment="false" DoNotApplyStrikeoutWithUnderline="false" DoNotApplyVertOffsetOfForward="false" DoNotApplyWhiteSpaceHeight="false" DoNotFormattingAtBeneathAnchor="false" DoNotHoldAnchorOfTable="false" ExtendLineheightToOffset="false" ExtendLineheightToParaBorderOffset="false" ExtendVertLimitToPageMargins="false" FixedUnderlineWidth="false" OverlapBothAllowOverlap="false" TreatQuotationAsLatin="false" UseInnerUnderline="false" UseLowercaseStrikeout="false" /></COMPATIBLEDOCUMENT></HEAD><BODY><SECDEF CharGrid="0" FirstBorder="false" FirstFill="false" LineGrid="0" OutlineShape="1" SpaceColumns="1134" TabStop="8000" TextDirection="0" TextVerticalWidthHead="0"><STARTNUMBER Equation="0" Figure="0" Page="0" PageStartsOn="Both" Table="0" /><HIDE Border="false" EmptyLine="false" Footer="false" Header="false" MasterPage="false" PageNumPos="false" /><PAGEDEF GutterType="LeftOnly" Height="84188" Landscape="0" Width="59528"><PAGEMARGIN Bottom="4252" Footer="4252" Gutter="0" Header="4252" Left="8504" Right="8504" Top="5668" /></PAGEDEF><FOOTNOTESHAPE><AUTONUMFORMAT SuffixChar=")" Superscript="false" Type="Digit" /><NOTELINE Length="5cm" Type="Solid" Width="0.12mm" /><NOTESPACING AboveLine="850" BelowLine="567" BetweenNotes="283" /><NOTENUMBERING NewNumber="1" Type="Continuous" /><NOTEPLACEMENT BeneathText="false" Place="EachColumn" /></FOOTNOTESHAPE><ENDNOTESHAPE><AUTONUMFORMAT SuffixChar=")" Superscript="false" Type="Digit" /><NOTELINE Length="14692344" Type="Solid" Width="0.12mm" /><NOTESPACING AboveLine="850" BelowLine="567" BetweenNotes="0" /><NOTENUMBERING NewNumber="1" Type="Continuous" /><NOTEPLACEMENT BeneathText="false" Place="EndOfDocument" /></ENDNOTESHAPE><PAGEBORDERFILL BorferFill="1" FillArea="Paper" FooterInside="false" HeaderInside="false" TextBorder="true" Type="Both"><PAGEOFFSET Bottom="1417" Left="1417" Right="1417" Top="1417" /></PAGEBORDERFILL><PAGEBORDERFILL BorferFill="1" FillArea="Paper" FooterInside="false" HeaderInside="false" TextBorder="true" Type="Even"><PAGEOFFSET Bottom="1417" Left="1417" Right="1417" Top="1417" /></PAGEBORDERFILL><PAGEBORDERFILL BorferFill="1" FillArea="Paper" FooterInside="false" HeaderInside="false" TextBorder="true" Type="Odd"><PAGEOFFSET Bottom="1417" Left="1417" Right="1417" Top="1417" /></PAGEBORDERFILL></SECDEF><SECTION Id="0">` + bindata_str;
}

const FOOTER = `</SECTION></BODY><TAIL></TAIL></HWPML>`;
const INTRO = `
`;

// Main Execution
const args = process.argv.slice(2);
let fIn = "README.md";
let fOut = "readme.hml";

if (args.length > 0) fIn = args[0];
if (args.length > 1) fOut = args[1];

console.log(`in:${fIn} -> out:${fOut}`);

if (!fs.existsSync(fIn)) {
    console.error(`Error: File ${fIn} not found.`);
    process.exit(1);
}

let ret = fs.readFileSync(fIn, 'utf8');

// Normalize replacements
ret = ret.replace(/\r/g, "");

// Start with newline
ret = "\n" + ret;

// 1. XML Escape Text Content (FIRST)
ret = xmlEscape(ret);

// Parse Metadata
let t = "No Title";
let a = "Unknown";
let d = "Unknown";

let t_match = ret.match(/(?<=title: ).*/);
if (t_match) t = t_match[0];
let a_match = ret.match(/(?<=author: ).*/);
if (a_match) a = a_match[0];
let d_match = ret.match(/(?<=date: ).*/);
if (d_match) d = d_match[0];

// Remove Front Matter
ret = ret.replace(/---[\s\S]*?---/, "");

// 2. Images: ![alt](path)
ret = ret.replace(/!\[(.*?)\]\((.*?)\)/g, (match, p1, p2) => process_image(p1, p2));

// 3. Code Blocks
// Pre-processing to Table
ret = ret.replace(/^```([\s\S]*?)```/gm, (match, p1) => process_code_block(p1));

// 4. Page Break / HR
ret = ret.replace(/\n---/g, `\n<P ParaShape="10" Style="0"><TEXT CharShape="0"><CHAR></CHAR></TEXT></P>`);

// 5. Links: [text](url) -> text (url)
ret = ret.replace(/\[(.*?)\]\((.*?)\)/g, '$1 ($2)');

// 6. Lists & Headers
ret = ret.replace(/\n[-+*] /g, "\n" + L1_START);
ret = ret.replace(/\n    [-+*] /g, "\n" + L2_START);
ret = ret.replace(/\n        [-+*] /g, "\n" + L3_START);
ret = ret.replace(/\n            [-+*] /g, "\n" + L4_START);
ret = ret.replace(/\n                [-+*] /g, "\n" + L5_START);
ret = ret.replace(/\n                    [-+*] /g, "\n" + L6_START);

// Ordered Lists
ret = ret.replace(/\n(\d+)\. /g, '\n<P ParaShape="12" Style="0"><TEXT CharShape="0"><CHAR>$1. ');
// Nested Ordered Lists
ret = ret.replace(/\n    (\d+)\. /g, '\n<P ParaShape="13" Style="0"><TEXT CharShape="0"><CHAR>$1. ');

// Task Lists
ret = ret.replace(/\[ \]/g, '☐');
ret = ret.replace(/\[x\]/ig, '☑');

// Blockquotes > text
ret = ret.replace(/\n(&gt;|>) (.*)/g, '\n<P ParaShape="11" Style="0"><TEXT CharShape="8"><CHAR>$2</CHAR></TEXT></P>');

// Tables (Native XML)
ret = ret.replace(/((?:^\|.*\|$\n?)+)/gm, (match, p1) => process_table_native(p1));

// Headers
ret = ret.replace(/\n# /g, "\n" + H1_START);
ret = ret.replace(/\n## /g, "\n" + H2_START);
ret = ret.replace(/\n### /g, "\n" + H3_START);
ret = ret.replace(/\n#### /g, "\n" + H4_START);
ret = ret.replace(/\n##### /g, "\n" + H4_START);
ret = ret.replace(/\n###### /g, "\n" + H4_START);

// 7. Text Wrapping
ret = ret.replace(/\n(?=[^<])/g, "\n" + TXT_START);

// Close Tags
try {
    ret = ret.replace(/(?<!<\/P>)(?<!>)\n/g, TXT_END + "\n");
} catch (e) {
    console.error("Warning: Node.js version might be too old for regex lookbehind.");
}

ret = ret.replace(new RegExp("^" + TXT_END), "");


// Process Inline
ret = ret.replace(/(<TEXT CharShape="0"><CHAR>)([\s\S]*?)(<\/CHAR>)/g, (match, p1, p2, p3) => p1 + process_inline(p2) + p3);

// Header
let bindata_str = "";
if (BIN_DATA_ENTRIES.length > 0) {
    bindata_str = `<BINDATALIST Count="${BIN_DATA_ENTRIES.length}">${BIN_DATA_ENTRIES.join('')}</BINDATALIST>`;
}
let header_content = get_header(bindata_str);

ret = header_content + INTRO + ret + FOOTER;

ret = ret.replace("___TITLE___", t);
ret = ret.replace("___AUTHOR___", a);
ret = ret.replace("___DATE___", d);

fs.writeFileSync(fOut, ret, 'utf8');
