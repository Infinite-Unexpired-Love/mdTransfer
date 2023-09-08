package DocxGen;

import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigDecimal;
import java.nio.ByteBuffer;
import java.util.ArrayList;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import md.Reader;

import javax.imageio.ImageIO;

public class DocxGenerater {
    private final File docx;
    private final Reader source;

    /**
     * 生成word/document.xml文件、word/media目录和word/_rels/document.xml.rels文件
     *
     * @param zipOutputStream
     */
    private void mainContentGen(ZipOutputStream zipOutputStream) throws IOException {
        var fileLines = source.getFileLines();
        var imgUrls = source.getImgURLs();
        var imgLines = source.getImgLines();
        var imgRefs = new ArrayList<String>();
        var imgAlas = new ArrayList<String>();
        try {
            //生成word/document.xml文件
            zipOutputStream.putNextEntry(new ZipEntry("word/document.xml"));
            //生成固定的文件头
            String header = """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
                                xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
                                xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
                                xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
                                xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
                                xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
                                xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
                                xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
                                xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
                                xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
                                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
                                xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
                                xmlns:o="urn:schemas-microsoft-com:office:office"
                                xmlns:oel="http://schemas.microsoft.com/office/2019/extlst"
                                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml"
                                xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
                                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                                xmlns:w10="urn:schemas-microsoft-com:office:word"
                                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                                xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                                xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                                xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
                                xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
                                xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
                                xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
                                xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
                                xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
                                xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
                                xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
                                xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                                mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">
                    <w:body>
                    """;
            String footer = """
                    <w:sectPr w:rsidR="00693BD7">
                                <w:pgSz w:w="11906" w:h="16838"/>
                                <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720"
                                         w:gutter="0"/>
                                <w:cols w:space="720"/>
                                <w:docGrid w:type="lines" w:linePitch="312"/>
                    </w:sectPr>
                    </w:body>
                    </w:document>
                    """;
            zipOutputStream.write(header.getBytes());
            //生成文件主体内容
            int imgNum = 0;
            for (int pos = 0; pos < fileLines.size(); pos++) {
                if ( imgNum==imgLines.size() || pos != imgLines.get(imgNum)) {
                    createParagraph(zipOutputStream, fileLines.get(pos), pos);
                } else {
                    createPicture(zipOutputStream, imgUrls.get(imgNum), imgNum, imgRefs);
                    imgNum++;
                }
            }
            //生成文件固定结尾
            zipOutputStream.write(footer.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建word/document.xml");
            //生成media目录
            createMedia(zipOutputStream, imgAlas);
            System.out.println("[+] 成功创建word/media目录");
            //生成word/_rels/document.xml.rels文件
            createRefs(zipOutputStream, imgRefs, imgAlas);
            System.out.println("[+] 成功创建word/_rels/document.xml.rels");
        }
        catch (IOException e){
            System.out.println("[-] 关键文件创建失败，docx文件已损坏");
            throw e;
        }
    }

    private void relatedGen(ZipOutputStream zipOutputStream) throws IOException {
        try {
            //生成docProps/app.xml
            String data1 = """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
                                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                        <Template>Normal.dotm</Template>
                        <TotalTime>0</TotalTime>
                        <Pages>1</Pages>
                        <Words>5</Words>
                        <Characters>32</Characters>
                        <Application>Microsoft Office Word</Application>
                        <DocSecurity>0</DocSecurity>
                        <Lines>1</Lines>
                        <Paragraphs>1</Paragraphs>
                        <ScaleCrop>false</ScaleCrop>
                        <Company></Company>
                        <LinksUpToDate>false</LinksUpToDate>
                        <CharactersWithSpaces>36</CharactersWithSpaces>
                        <SharedDoc>false</SharedDoc>
                        <HyperlinksChanged>false</HyperlinksChanged>
                        <AppVersion>16.0000</AppVersion>
                    </Properties>
                    """;
            zipOutputStream.putNextEntry(new ZipEntry("docProps/app.xml"));
            zipOutputStream.write(data1.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建docProps/app.xml");
            //生成docProps/core.xml
            String data2 = """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                                       xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
                                       xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                        <dc:creator>Infinite-Unexpired-Love</dc:creator>
                        <cp:lastModifiedBy>Infinite-Unexpired-Love</cp:lastModifiedBy>
                        <cp:revision>5</cp:revision>
                        <dcterms:created xsi:type="dcterms:W3CDTF">2023-09-06T15:41:00Z</dcterms:created>
                        <dcterms:modified xsi:type="dcterms:W3CDTF">2023-09-07T08:34:00Z</dcterms:modified>
                    </cp:coreProperties>
                    """;
            zipOutputStream.putNextEntry(new ZipEntry("docProps/core.xml"));
            zipOutputStream.write(data2.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建docProps/core.xml");
            //生成docProps/custom.xml
            String data3 = """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                        <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="KSOProductBuildVer">
                            <vt:lpwstr>2052-11.1.0.14309</vt:lpwstr>
                        </property>
                        <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="ICV">
                            <vt:lpwstr>2AA82F9CD5484CACBE23D7CE24647AAD_12</vt:lpwstr>
                        </property>
                    </Properties>
                    """;
            zipOutputStream.putNextEntry(new ZipEntry("docProps/custom.xml"));
            zipOutputStream.write(data3.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建docProps/custom.xml");
            //生成word/theme/theme1.xml
            String data4 = """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题​​">
                        <a:themeElements>
                            <a:clrScheme name="Office">
                                <a:dk1>
                                    <a:sysClr val="windowText" lastClr="000000"/>
                                </a:dk1>
                                <a:lt1>
                                    <a:sysClr val="window" lastClr="FFFFFF"/>
                                </a:lt1>
                                <a:dk2>
                                    <a:srgbClr val="44546A"/>
                                </a:dk2>
                                <a:lt2>
                                    <a:srgbClr val="E7E6E6"/>
                                </a:lt2>
                                <a:accent1>
                                    <a:srgbClr val="4472C4"/>
                                </a:accent1>
                                <a:accent2>
                                    <a:srgbClr val="ED7D31"/>
                                </a:accent2>
                                <a:accent3>
                                    <a:srgbClr val="A5A5A5"/>
                                </a:accent3>
                                <a:accent4>
                                    <a:srgbClr val="FFC000"/>
                                </a:accent4>
                                <a:accent5>
                                    <a:srgbClr val="5B9BD5"/>
                                </a:accent5>
                                <a:accent6>
                                    <a:srgbClr val="70AD47"/>
                                </a:accent6>
                                <a:hlink>
                                    <a:srgbClr val="0563C1"/>
                                </a:hlink>
                                <a:folHlink>
                                    <a:srgbClr val="954F72"/>
                                </a:folHlink>
                            </a:clrScheme>
                            <a:fontScheme name="Office">
                                <a:majorFont>
                                    <a:latin typeface="等线 Light"/>
                                    <a:ea typeface=""/>
                                    <a:cs typeface=""/>
                                    <a:font script="Jpan" typeface="游ゴシック Light"/>
                                    <a:font script="Hang" typeface="맑은 고딕"/>
                                    <a:font script="Hans" typeface="等线 Light"/>
                                    <a:font script="Hant" typeface="新細明體"/>
                                    <a:font script="Arab" typeface="Times New Roman"/>
                                    <a:font script="Hebr" typeface="Times New Roman"/>
                                    <a:font script="Thai" typeface="Angsana New"/>
                                    <a:font script="Ethi" typeface="Nyala"/>
                                    <a:font script="Beng" typeface="Vrinda"/>
                                    <a:font script="Gujr" typeface="Shruti"/>
                                    <a:font script="Khmr" typeface="MoolBoran"/>
                                    <a:font script="Knda" typeface="Tunga"/>
                                    <a:font script="Guru" typeface="Raavi"/>
                                    <a:font script="Cans" typeface="Euphemia"/>
                                    <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                                    <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                                    <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                                    <a:font script="Thaa" typeface="MV Boli"/>
                                    <a:font script="Deva" typeface="Mangal"/>
                                    <a:font script="Telu" typeface="Gautami"/>
                                    <a:font script="Taml" typeface="Latha"/>
                                    <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                                    <a:font script="Orya" typeface="Kalinga"/>
                                    <a:font script="Mlym" typeface="Kartika"/>
                                    <a:font script="Laoo" typeface="DokChampa"/>
                                    <a:font script="Sinh" typeface="Iskoola Pota"/>
                                    <a:font script="Mong" typeface="Mongolian Baiti"/>
                                    <a:font script="Viet" typeface="Times New Roman"/>
                                    <a:font script="Uigh" typeface="Microsoft Uighur"/>
                                    <a:font script="Geor" typeface="Sylfaen"/>
                                </a:majorFont>
                                <a:minorFont>
                                    <a:latin typeface="等线"/>
                                    <a:ea typeface=""/>
                                    <a:cs typeface=""/>
                                    <a:font script="Jpan" typeface="游明朝"/>
                                    <a:font script="Hang" typeface="맑은 고딕"/>
                                    <a:font script="Hans" typeface="等线"/>
                                    <a:font script="Hant" typeface="新細明體"/>
                                    <a:font script="Arab" typeface="Arial"/>
                                    <a:font script="Hebr" typeface="Arial"/>
                                    <a:font script="Thai" typeface="Cordia New"/>
                                    <a:font script="Ethi" typeface="Nyala"/>
                                    <a:font script="Beng" typeface="Vrinda"/>
                                    <a:font script="Gujr" typeface="Shruti"/>
                                    <a:font script="Khmr" typeface="DaunPenh"/>
                                    <a:font script="Knda" typeface="Tunga"/>
                                    <a:font script="Guru" typeface="Raavi"/>
                                    <a:font script="Cans" typeface="Euphemia"/>
                                    <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                                    <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                                    <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                                    <a:font script="Thaa" typeface="MV Boli"/>
                                    <a:font script="Deva" typeface="Mangal"/>
                                    <a:font script="Telu" typeface="Gautami"/>
                                    <a:font script="Taml" typeface="Latha"/>
                                    <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                                    <a:font script="Orya" typeface="Kalinga"/>
                                    <a:font script="Mlym" typeface="Kartika"/>
                                    <a:font script="Laoo" typeface="DokChampa"/>
                                    <a:font script="Sinh" typeface="Iskoola Pota"/>
                                    <a:font script="Mong" typeface="Mongolian Baiti"/>
                                    <a:font script="Viet" typeface="Arial"/>
                                    <a:font script="Uigh" typeface="Microsoft Uighur"/>
                                    <a:font script="Geor" typeface="Sylfaen"/>
                                </a:minorFont>
                            </a:fontScheme>
                            <a:fmtScheme name="Office">
                                <a:fillStyleLst>
                                    <a:solidFill>
                                        <a:schemeClr val="phClr"/>
                                    </a:solidFill>
                                    <a:gradFill rotWithShape="1">
                                        <a:gsLst>
                                            <a:gs pos="0">
                                                <a:schemeClr val="phClr">
                                                    <a:lumMod val="110000"/>
                                                    <a:satMod val="105000"/>
                                                    <a:tint val="67000"/>
                                                </a:schemeClr>
                                            </a:gs>
                                            <a:gs pos="50000">
                                                <a:schemeClr val="phClr">
                                                    <a:lumMod val="105000"/>
                                                    <a:satMod val="103000"/>
                                                    <a:tint val="73000"/>
                                                </a:schemeClr>
                                            </a:gs>
                                            <a:gs pos="100000">
                                                <a:schemeClr val="phClr">
                                                    <a:lumMod val="105000"/>
                                                    <a:satMod val="109000"/>
                                                    <a:tint val="81000"/>
                                                </a:schemeClr>
                                            </a:gs>
                                        </a:gsLst>
                                        <a:lin ang="5400000" scaled="0"/>
                                    </a:gradFill>
                                    <a:gradFill rotWithShape="1">
                                        <a:gsLst>
                                            <a:gs pos="0">
                                                <a:schemeClr val="phClr">
                                                    <a:satMod val="103000"/>
                                                    <a:lumMod val="102000"/>
                                                    <a:tint val="94000"/>
                                                </a:schemeClr>
                                            </a:gs>
                                            <a:gs pos="50000">
                                                <a:schemeClr val="phClr">
                                                    <a:satMod val="110000"/>
                                                    <a:lumMod val="100000"/>
                                                    <a:shade val="100000"/>
                                                </a:schemeClr>
                                            </a:gs>
                                            <a:gs pos="100000">
                                                <a:schemeClr val="phClr">
                                                    <a:lumMod val="99000"/>
                                                    <a:satMod val="120000"/>
                                                    <a:shade val="78000"/>
                                                </a:schemeClr>
                                            </a:gs>
                                        </a:gsLst>
                                        <a:lin ang="5400000" scaled="0"/>
                                    </a:gradFill>
                                </a:fillStyleLst>
                                <a:lnStyleLst>
                                    <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
                                        <a:solidFill>
                                            <a:schemeClr val="phClr"/>
                                        </a:solidFill>
                                        <a:prstDash val="solid"/>
                                        <a:miter lim="800000"/>
                                    </a:ln>
                                    <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
                                        <a:solidFill>
                                            <a:schemeClr val="phClr"/>
                                        </a:solidFill>
                                        <a:prstDash val="solid"/>
                                        <a:miter lim="800000"/>
                                    </a:ln>
                                    <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
                                        <a:solidFill>
                                            <a:schemeClr val="phClr"/>
                                        </a:solidFill>
                                        <a:prstDash val="solid"/>
                                        <a:miter lim="800000"/>
                                    </a:ln>
                                </a:lnStyleLst>
                                <a:effectStyleLst>
                                    <a:effectStyle>
                                        <a:effectLst/>
                                    </a:effectStyle>
                                    <a:effectStyle>
                                        <a:effectLst/>
                                    </a:effectStyle>
                                    <a:effectStyle>
                                        <a:effectLst>
                                            <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
                                                <a:srgbClr val="000000">
                                                    <a:alpha val="63000"/>
                                                </a:srgbClr>
                                            </a:outerShdw>
                                        </a:effectLst>
                                    </a:effectStyle>
                                </a:effectStyleLst>
                                <a:bgFillStyleLst>
                                    <a:solidFill>
                                        <a:schemeClr val="phClr"/>
                                    </a:solidFill>
                                    <a:solidFill>
                                        <a:schemeClr val="phClr">
                                            <a:tint val="95000"/>
                                            <a:satMod val="170000"/>
                                        </a:schemeClr>
                                    </a:solidFill>
                                    <a:gradFill rotWithShape="1">
                                        <a:gsLst>
                                            <a:gs pos="0">
                                                <a:schemeClr val="phClr">
                                                    <a:tint val="93000"/>
                                                    <a:satMod val="150000"/>
                                                    <a:shade val="98000"/>
                                                    <a:lumMod val="102000"/>
                                                </a:schemeClr>
                                            </a:gs>
                                            <a:gs pos="50000">
                                                <a:schemeClr val="phClr">
                                                    <a:tint val="98000"/>
                                                    <a:satMod val="130000"/>
                                                    <a:shade val="90000"/>
                                                    <a:lumMod val="103000"/>
                                                </a:schemeClr>
                                            </a:gs>
                                            <a:gs pos="100000">
                                                <a:schemeClr val="phClr">
                                                    <a:shade val="63000"/>
                                                    <a:satMod val="120000"/>
                                                </a:schemeClr>
                                            </a:gs>
                                        </a:gsLst>
                                        <a:lin ang="5400000" scaled="0"/>
                                    </a:gradFill>
                                </a:bgFillStyleLst>
                            </a:fmtScheme>
                        </a:themeElements>
                        <a:objectDefaults/>
                        <a:extraClrSchemeLst/>
                    </a:theme>
                    """;
            zipOutputStream.putNextEntry(new ZipEntry("word/theme/theme1.xml"));
            zipOutputStream.write(data4.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建word/theme/theme1.xml");
            //生成word/fontTable.xml
            String data5 = """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <w:fonts xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                             xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                             xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                             xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                             xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
                             xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
                             xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
                             xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
                             xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
                             mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">
                        <w:font w:name="等线">
                            <w:altName w:val="DengXian"/>
                            <w:panose1 w:val="02010600030101010101"/>
                            <w:charset w:val="86"/>
                            <w:family w:val="auto"/>
                            <w:pitch w:val="variable"/>
                            <w:sig w:usb0="A00002BF" w:usb1="38CF7CFA" w:usb2="00000016" w:usb3="00000000" w:csb0="0004000F"
                                   w:csb1="00000000"/>
                        </w:font>
                        <w:font w:name="Times New Roman">
                            <w:panose1 w:val="02020603050405020304"/>
                            <w:charset w:val="00"/>
                            <w:family w:val="roman"/>
                            <w:pitch w:val="variable"/>
                            <w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF"
                                   w:csb1="00000000"/>
                        </w:font>
                        <w:font w:name="等线 Light">
                            <w:panose1 w:val="02010600030101010101"/>
                            <w:charset w:val="86"/>
                            <w:family w:val="auto"/>
                            <w:pitch w:val="variable"/>
                            <w:sig w:usb0="A00002BF" w:usb1="38CF7CFA" w:usb2="00000016" w:usb3="00000000" w:csb0="0004000F"
                                   w:csb1="00000000"/>
                        </w:font>
                    </w:fonts>
                    """;
            zipOutputStream.putNextEntry(new ZipEntry("word/fontTable.xml"));
            zipOutputStream.write(data5.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建word/fontTable.xml");
            //生成word/settings.xml
            String data6 = """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                xmlns:o="urn:schemas-microsoft-com:office:office"
                                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml"
                                xmlns:w10="urn:schemas-microsoft-com:office:word"
                                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                                xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                                xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                                xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
                                xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
                                xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
                                xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
                                xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
                                xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
                                mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">
                        <w:zoom w:percent="100"/>
                        <w:bordersDoNotSurroundHeader/>
                        <w:bordersDoNotSurroundFooter/>
                        <w:defaultTabStop w:val="420"/>
                        <w:drawingGridVerticalSpacing w:val="156"/>
                        <w:displayHorizontalDrawingGridEvery w:val="0"/>
                        <w:displayVerticalDrawingGridEvery w:val="2"/>
                        <w:characterSpacingControl w:val="compressPunctuation"/>
                        <w:compat>
                            <w:spaceForUL/>
                            <w:balanceSingleByteDoubleByteWidth/>
                            <w:doNotLeaveBackslashAlone/>
                            <w:ulTrailSpace/>
                            <w:doNotExpandShiftReturn/>
                            <w:adjustLineHeightInTable/>
                            <w:useFELayout/>
                            <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
                            <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification"
                                             w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
                            <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
                            <w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
                            <w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word"
                                             w:val="1"/>
                            <w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word"
                                             w:val="1"/>
                        </w:compat>
                        <w:docVars>
                            <w:docVar w:name="commondata" w:val="eyJoZGlkIjoiMzFlZTE2MTNiNTc0NDE0MDMyOTBmZDBkYTRiMWQyOTQifQ=="/>
                        </w:docVars>
                        <w:rsids>
                            <w:rsidRoot w:val="005E02E1"/>
                            <w:rsid w:val="00451324"/>
                            <w:rsid w:val="005E02E1"/>
                            <w:rsid w:val="00662722"/>
                            <w:rsid w:val="00693BD7"/>
                            <w:rsid w:val="00D94B4D"/>
                            <w:rsid w:val="00E22084"/>
                            <w:rsid w:val="00F34C9C"/>
                            <w:rsid w:val="049F508F"/>
                        </w:rsids>
                        <m:mathPr>
                            <m:mathFont m:val="Cambria Math"/>
                            <m:brkBin m:val="before"/>
                            <m:brkBinSub m:val="--"/>
                            <m:smallFrac m:val="0"/>
                            <m:dispDef/>
                            <m:lMargin m:val="0"/>
                            <m:rMargin m:val="0"/>
                            <m:defJc m:val="centerGroup"/>
                            <m:wrapIndent m:val="1440"/>
                            <m:intLim m:val="subSup"/>
                            <m:naryLim m:val="undOvr"/>
                        </m:mathPr>
                        <w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/>
                        <w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2"
                                            w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6"
                                            w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>
                        <w:shapeDefaults>
                            <o:shapedefaults v:ext="edit" spidmax="1026"/>
                            <o:shapelayout v:ext="edit">
                                <o:idmap v:ext="edit" data="1"/>
                            </o:shapelayout>
                        </w:shapeDefaults>
                        <w:decimalSymbol w:val="."/>
                        <w:listSeparator w:val=","/>
                        <w14:docId w14:val="3034704B"/>
                        <w15:docId w15:val="{25F91FF6-8FA7-4E4C-B195-4DDB4E22B3B3}"/>
                    </w:settings>
                    """;
            zipOutputStream.putNextEntry(new ZipEntry("word/settings.xml"));
            zipOutputStream.write(data6.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建word/settings.xml");
            //生成word/styles.xml
            String data7 = """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                              xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                              xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                              xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                              xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
                              xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
                              xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
                              xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
                              xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
                              mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">
                        <w:docDefaults>
                            <w:rPrDefault>
                                <w:rPr>
                                    <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi"
                                              w:cstheme="minorBidi"/>
                                    <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
                                </w:rPr>
                            </w:rPrDefault>
                            <w:pPrDefault/>
                        </w:docDefaults>
                        <w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0"
                                        w:defQFormat="0" w:count="376">
                            <w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
                            <w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
                            <w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
                            <w:lsdException w:name="heading 3" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
                            <w:lsdException w:name="heading 4" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
                            <w:lsdException w:name="heading 5" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
                            <w:lsdException w:name="heading 6" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
                            <w:lsdException w:name="heading 7" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
                            <w:lsdException w:name="heading 8" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
                            <w:lsdException w:name="heading 9" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
                            <w:lsdException w:name="index 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="index 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="index 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="index 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="index 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="index 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="index 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="index 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="index 9" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="toc 1" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="toc 2" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="toc 3" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="toc 4" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="toc 5" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="toc 6" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="toc 7" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="toc 8" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="toc 9" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Normal Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="footnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="annotation text" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="header" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="footer" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="index heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="caption" w:semiHidden="1" w:uiPriority="35" w:unhideWhenUsed="1" w:qFormat="1"/>
                            <w:lsdException w:name="table of figures" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="envelope address" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="envelope return" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="footnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="annotation reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="line number" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="page number" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="endnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="endnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="table of authorities" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="macro" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="toa heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Bullet" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Number" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Bullet 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Bullet 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Bullet 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Bullet 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Number 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Number 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Number 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Number 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/>
                            <w:lsdException w:name="Closing" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Default Paragraph Font" w:semiHidden="1" w:uiPriority="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Body Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Body Text Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Continue" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Continue 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Continue 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Continue 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Continue 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Message Header" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Subtitle" w:uiPriority="11" w:qFormat="1"/>
                            <w:lsdException w:name="Salutation" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Date" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Body Text First Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Body Text First Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Note Heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Body Text 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Body Text 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Body Text Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Body Text Indent 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Block Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="FollowedHyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Strong" w:uiPriority="22" w:qFormat="1"/>
                            <w:lsdException w:name="Emphasis" w:uiPriority="20" w:qFormat="1"/>
                            <w:lsdException w:name="Document Map" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Plain Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="E-mail Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Top of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Bottom of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Normal (Web)" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Acronym" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Address" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Cite" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Code" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Definition" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Keyboard" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Preformatted" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Sample" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Typewriter" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="HTML Variable" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Normal Table" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="annotation subject" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="No List" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Outline List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Outline List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Outline List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Simple 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Simple 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Simple 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Classic 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Classic 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Classic 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Classic 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Colorful 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Colorful 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Colorful 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Columns 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Columns 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Columns 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Columns 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Columns 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Grid 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Grid 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Grid 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Grid 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Grid 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Grid 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Grid 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Grid 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table List 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table List 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table List 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table 3D effects 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table 3D effects 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table 3D effects 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Contemporary" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Elegant" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Professional" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Subtle 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Subtle 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Web 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Web 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Web 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Balloon Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Table Grid" w:uiPriority="39"/>
                            <w:lsdException w:name="Table Theme" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Placeholder Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="No Spacing" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Light Shading" w:uiPriority="60"/>
                            <w:lsdException w:name="Light List" w:uiPriority="61"/>
                            <w:lsdException w:name="Light Grid" w:uiPriority="62"/>
                            <w:lsdException w:name="Medium Shading 1" w:uiPriority="63"/>
                            <w:lsdException w:name="Medium Shading 2" w:uiPriority="64"/>
                            <w:lsdException w:name="Medium List 1" w:uiPriority="65"/>
                            <w:lsdException w:name="Medium List 2" w:uiPriority="66"/>
                            <w:lsdException w:name="Medium Grid 1" w:uiPriority="67"/>
                            <w:lsdException w:name="Medium Grid 2" w:uiPriority="68"/>
                            <w:lsdException w:name="Medium Grid 3" w:uiPriority="69"/>
                            <w:lsdException w:name="Dark List" w:uiPriority="70"/>
                            <w:lsdException w:name="Colorful Shading" w:uiPriority="71"/>
                            <w:lsdException w:name="Colorful List" w:uiPriority="72"/>
                            <w:lsdException w:name="Colorful Grid" w:uiPriority="73"/>
                            <w:lsdException w:name="Light Shading Accent 1" w:uiPriority="60"/>
                            <w:lsdException w:name="Light List Accent 1" w:uiPriority="61"/>
                            <w:lsdException w:name="Light Grid Accent 1" w:uiPriority="62"/>
                            <w:lsdException w:name="Medium Shading 1 Accent 1" w:uiPriority="63"/>
                            <w:lsdException w:name="Medium Shading 2 Accent 1" w:uiPriority="64"/>
                            <w:lsdException w:name="Medium List 1 Accent 1" w:uiPriority="65"/>
                            <w:lsdException w:name="Revision" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="List Paragraph" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Quote" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Intense Quote" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Medium List 2 Accent 1" w:uiPriority="66"/>
                            <w:lsdException w:name="Medium Grid 1 Accent 1" w:uiPriority="67"/>
                            <w:lsdException w:name="Medium Grid 2 Accent 1" w:uiPriority="68"/>
                            <w:lsdException w:name="Medium Grid 3 Accent 1" w:uiPriority="69"/>
                            <w:lsdException w:name="Dark List Accent 1" w:uiPriority="70"/>
                            <w:lsdException w:name="Colorful Shading Accent 1" w:uiPriority="71"/>
                            <w:lsdException w:name="Colorful List Accent 1" w:uiPriority="72"/>
                            <w:lsdException w:name="Colorful Grid Accent 1" w:uiPriority="73"/>
                            <w:lsdException w:name="Light Shading Accent 2" w:uiPriority="60"/>
                            <w:lsdException w:name="Light List Accent 2" w:uiPriority="61"/>
                            <w:lsdException w:name="Light Grid Accent 2" w:uiPriority="62"/>
                            <w:lsdException w:name="Medium Shading 1 Accent 2" w:uiPriority="63"/>
                            <w:lsdException w:name="Medium Shading 2 Accent 2" w:uiPriority="64"/>
                            <w:lsdException w:name="Medium List 1 Accent 2" w:uiPriority="65"/>
                            <w:lsdException w:name="Medium List 2 Accent 2" w:uiPriority="66"/>
                            <w:lsdException w:name="Medium Grid 1 Accent 2" w:uiPriority="67"/>
                            <w:lsdException w:name="Medium Grid 2 Accent 2" w:uiPriority="68"/>
                            <w:lsdException w:name="Medium Grid 3 Accent 2" w:uiPriority="69"/>
                            <w:lsdException w:name="Dark List Accent 2" w:uiPriority="70"/>
                            <w:lsdException w:name="Colorful Shading Accent 2" w:uiPriority="71"/>
                            <w:lsdException w:name="Colorful List Accent 2" w:uiPriority="72"/>
                            <w:lsdException w:name="Colorful Grid Accent 2" w:uiPriority="73"/>
                            <w:lsdException w:name="Light Shading Accent 3" w:uiPriority="60"/>
                            <w:lsdException w:name="Light List Accent 3" w:uiPriority="61"/>
                            <w:lsdException w:name="Light Grid Accent 3" w:uiPriority="62"/>
                            <w:lsdException w:name="Medium Shading 1 Accent 3" w:uiPriority="63"/>
                            <w:lsdException w:name="Medium Shading 2 Accent 3" w:uiPriority="64"/>
                            <w:lsdException w:name="Medium List 1 Accent 3" w:uiPriority="65"/>
                            <w:lsdException w:name="Medium List 2 Accent 3" w:uiPriority="66"/>
                            <w:lsdException w:name="Medium Grid 1 Accent 3" w:uiPriority="67"/>
                            <w:lsdException w:name="Medium Grid 2 Accent 3" w:uiPriority="68"/>
                            <w:lsdException w:name="Medium Grid 3 Accent 3" w:uiPriority="69"/>
                            <w:lsdException w:name="Dark List Accent 3" w:uiPriority="70"/>
                            <w:lsdException w:name="Colorful Shading Accent 3" w:uiPriority="71"/>
                            <w:lsdException w:name="Colorful List Accent 3" w:uiPriority="72"/>
                            <w:lsdException w:name="Colorful Grid Accent 3" w:uiPriority="73"/>
                            <w:lsdException w:name="Light Shading Accent 4" w:uiPriority="60"/>
                            <w:lsdException w:name="Light List Accent 4" w:uiPriority="61"/>
                            <w:lsdException w:name="Light Grid Accent 4" w:uiPriority="62"/>
                            <w:lsdException w:name="Medium Shading 1 Accent 4" w:uiPriority="63"/>
                            <w:lsdException w:name="Medium Shading 2 Accent 4" w:uiPriority="64"/>
                            <w:lsdException w:name="Medium List 1 Accent 4" w:uiPriority="65"/>
                            <w:lsdException w:name="Medium List 2 Accent 4" w:uiPriority="66"/>
                            <w:lsdException w:name="Medium Grid 1 Accent 4" w:uiPriority="67"/>
                            <w:lsdException w:name="Medium Grid 2 Accent 4" w:uiPriority="68"/>
                            <w:lsdException w:name="Medium Grid 3 Accent 4" w:uiPriority="69"/>
                            <w:lsdException w:name="Dark List Accent 4" w:uiPriority="70"/>
                            <w:lsdException w:name="Colorful Shading Accent 4" w:uiPriority="71"/>
                            <w:lsdException w:name="Colorful List Accent 4" w:uiPriority="72"/>
                            <w:lsdException w:name="Colorful Grid Accent 4" w:uiPriority="73"/>
                            <w:lsdException w:name="Light Shading Accent 5" w:uiPriority="60"/>
                            <w:lsdException w:name="Light List Accent 5" w:uiPriority="61"/>
                            <w:lsdException w:name="Light Grid Accent 5" w:uiPriority="62"/>
                            <w:lsdException w:name="Medium Shading 1 Accent 5" w:uiPriority="63"/>
                            <w:lsdException w:name="Medium Shading 2 Accent 5" w:uiPriority="64"/>
                            <w:lsdException w:name="Medium List 1 Accent 5" w:uiPriority="65"/>
                            <w:lsdException w:name="Medium List 2 Accent 5" w:uiPriority="66"/>
                            <w:lsdException w:name="Medium Grid 1 Accent 5" w:uiPriority="67"/>
                            <w:lsdException w:name="Medium Grid 2 Accent 5" w:uiPriority="68"/>
                            <w:lsdException w:name="Medium Grid 3 Accent 5" w:uiPriority="69"/>
                            <w:lsdException w:name="Dark List Accent 5" w:uiPriority="70"/>
                            <w:lsdException w:name="Colorful Shading Accent 5" w:uiPriority="71"/>
                            <w:lsdException w:name="Colorful List Accent 5" w:uiPriority="72"/>
                            <w:lsdException w:name="Colorful Grid Accent 5" w:uiPriority="73"/>
                            <w:lsdException w:name="Light Shading Accent 6" w:uiPriority="60"/>
                            <w:lsdException w:name="Light List Accent 6" w:uiPriority="61"/>
                            <w:lsdException w:name="Light Grid Accent 6" w:uiPriority="62"/>
                            <w:lsdException w:name="Medium Shading 1 Accent 6" w:uiPriority="63"/>
                            <w:lsdException w:name="Medium Shading 2 Accent 6" w:uiPriority="64"/>
                            <w:lsdException w:name="Medium List 1 Accent 6" w:uiPriority="65"/>
                            <w:lsdException w:name="Medium List 2 Accent 6" w:uiPriority="66"/>
                            <w:lsdException w:name="Medium Grid 1 Accent 6" w:uiPriority="67"/>
                            <w:lsdException w:name="Medium Grid 2 Accent 6" w:uiPriority="68"/>
                            <w:lsdException w:name="Medium Grid 3 Accent 6" w:uiPriority="69"/>
                            <w:lsdException w:name="Dark List Accent 6" w:uiPriority="70"/>
                            <w:lsdException w:name="Colorful Shading Accent 6" w:uiPriority="71"/>
                            <w:lsdException w:name="Colorful List Accent 6" w:uiPriority="72"/>
                            <w:lsdException w:name="Colorful Grid Accent 6" w:uiPriority="73"/>
                            <w:lsdException w:name="Subtle Emphasis" w:uiPriority="19" w:qFormat="1"/>
                            <w:lsdException w:name="Intense Emphasis" w:uiPriority="21" w:qFormat="1"/>
                            <w:lsdException w:name="Subtle Reference" w:uiPriority="31" w:qFormat="1"/>
                            <w:lsdException w:name="Intense Reference" w:uiPriority="32" w:qFormat="1"/>
                            <w:lsdException w:name="Book Title" w:uiPriority="33" w:qFormat="1"/>
                            <w:lsdException w:name="Bibliography" w:semiHidden="1" w:uiPriority="37" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="TOC Heading" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1" w:qFormat="1"/>
                            <w:lsdException w:name="Plain Table 1" w:uiPriority="41"/>
                            <w:lsdException w:name="Plain Table 2" w:uiPriority="42"/>
                            <w:lsdException w:name="Plain Table 3" w:uiPriority="43"/>
                            <w:lsdException w:name="Plain Table 4" w:uiPriority="44"/>
                            <w:lsdException w:name="Plain Table 5" w:uiPriority="45"/>
                            <w:lsdException w:name="Grid Table Light" w:uiPriority="40"/>
                            <w:lsdException w:name="Grid Table 1 Light" w:uiPriority="46"/>
                            <w:lsdException w:name="Grid Table 2" w:uiPriority="47"/>
                            <w:lsdException w:name="Grid Table 3" w:uiPriority="48"/>
                            <w:lsdException w:name="Grid Table 4" w:uiPriority="49"/>
                            <w:lsdException w:name="Grid Table 5 Dark" w:uiPriority="50"/>
                            <w:lsdException w:name="Grid Table 6 Colorful" w:uiPriority="51"/>
                            <w:lsdException w:name="Grid Table 7 Colorful" w:uiPriority="52"/>
                            <w:lsdException w:name="Grid Table 1 Light Accent 1" w:uiPriority="46"/>
                            <w:lsdException w:name="Grid Table 2 Accent 1" w:uiPriority="47"/>
                            <w:lsdException w:name="Grid Table 3 Accent 1" w:uiPriority="48"/>
                            <w:lsdException w:name="Grid Table 4 Accent 1" w:uiPriority="49"/>
                            <w:lsdException w:name="Grid Table 5 Dark Accent 1" w:uiPriority="50"/>
                            <w:lsdException w:name="Grid Table 6 Colorful Accent 1" w:uiPriority="51"/>
                            <w:lsdException w:name="Grid Table 7 Colorful Accent 1" w:uiPriority="52"/>
                            <w:lsdException w:name="Grid Table 1 Light Accent 2" w:uiPriority="46"/>
                            <w:lsdException w:name="Grid Table 2 Accent 2" w:uiPriority="47"/>
                            <w:lsdException w:name="Grid Table 3 Accent 2" w:uiPriority="48"/>
                            <w:lsdException w:name="Grid Table 4 Accent 2" w:uiPriority="49"/>
                            <w:lsdException w:name="Grid Table 5 Dark Accent 2" w:uiPriority="50"/>
                            <w:lsdException w:name="Grid Table 6 Colorful Accent 2" w:uiPriority="51"/>
                            <w:lsdException w:name="Grid Table 7 Colorful Accent 2" w:uiPriority="52"/>
                            <w:lsdException w:name="Grid Table 1 Light Accent 3" w:uiPriority="46"/>
                            <w:lsdException w:name="Grid Table 2 Accent 3" w:uiPriority="47"/>
                            <w:lsdException w:name="Grid Table 3 Accent 3" w:uiPriority="48"/>
                            <w:lsdException w:name="Grid Table 4 Accent 3" w:uiPriority="49"/>
                            <w:lsdException w:name="Grid Table 5 Dark Accent 3" w:uiPriority="50"/>
                            <w:lsdException w:name="Grid Table 6 Colorful Accent 3" w:uiPriority="51"/>
                            <w:lsdException w:name="Grid Table 7 Colorful Accent 3" w:uiPriority="52"/>
                            <w:lsdException w:name="Grid Table 1 Light Accent 4" w:uiPriority="46"/>
                            <w:lsdException w:name="Grid Table 2 Accent 4" w:uiPriority="47"/>
                            <w:lsdException w:name="Grid Table 3 Accent 4" w:uiPriority="48"/>
                            <w:lsdException w:name="Grid Table 4 Accent 4" w:uiPriority="49"/>
                            <w:lsdException w:name="Grid Table 5 Dark Accent 4" w:uiPriority="50"/>
                            <w:lsdException w:name="Grid Table 6 Colorful Accent 4" w:uiPriority="51"/>
                            <w:lsdException w:name="Grid Table 7 Colorful Accent 4" w:uiPriority="52"/>
                            <w:lsdException w:name="Grid Table 1 Light Accent 5" w:uiPriority="46"/>
                            <w:lsdException w:name="Grid Table 2 Accent 5" w:uiPriority="47"/>
                            <w:lsdException w:name="Grid Table 3 Accent 5" w:uiPriority="48"/>
                            <w:lsdException w:name="Grid Table 4 Accent 5" w:uiPriority="49"/>
                            <w:lsdException w:name="Grid Table 5 Dark Accent 5" w:uiPriority="50"/>
                            <w:lsdException w:name="Grid Table 6 Colorful Accent 5" w:uiPriority="51"/>
                            <w:lsdException w:name="Grid Table 7 Colorful Accent 5" w:uiPriority="52"/>
                            <w:lsdException w:name="Grid Table 1 Light Accent 6" w:uiPriority="46"/>
                            <w:lsdException w:name="Grid Table 2 Accent 6" w:uiPriority="47"/>
                            <w:lsdException w:name="Grid Table 3 Accent 6" w:uiPriority="48"/>
                            <w:lsdException w:name="Grid Table 4 Accent 6" w:uiPriority="49"/>
                            <w:lsdException w:name="Grid Table 5 Dark Accent 6" w:uiPriority="50"/>
                            <w:lsdException w:name="Grid Table 6 Colorful Accent 6" w:uiPriority="51"/>
                            <w:lsdException w:name="Grid Table 7 Colorful Accent 6" w:uiPriority="52"/>
                            <w:lsdException w:name="List Table 1 Light" w:uiPriority="46"/>
                            <w:lsdException w:name="List Table 2" w:uiPriority="47"/>
                            <w:lsdException w:name="List Table 3" w:uiPriority="48"/>
                            <w:lsdException w:name="List Table 4" w:uiPriority="49"/>
                            <w:lsdException w:name="List Table 5 Dark" w:uiPriority="50"/>
                            <w:lsdException w:name="List Table 6 Colorful" w:uiPriority="51"/>
                            <w:lsdException w:name="List Table 7 Colorful" w:uiPriority="52"/>
                            <w:lsdException w:name="List Table 1 Light Accent 1" w:uiPriority="46"/>
                            <w:lsdException w:name="List Table 2 Accent 1" w:uiPriority="47"/>
                            <w:lsdException w:name="List Table 3 Accent 1" w:uiPriority="48"/>
                            <w:lsdException w:name="List Table 4 Accent 1" w:uiPriority="49"/>
                            <w:lsdException w:name="List Table 5 Dark Accent 1" w:uiPriority="50"/>
                            <w:lsdException w:name="List Table 6 Colorful Accent 1" w:uiPriority="51"/>
                            <w:lsdException w:name="List Table 7 Colorful Accent 1" w:uiPriority="52"/>
                            <w:lsdException w:name="List Table 1 Light Accent 2" w:uiPriority="46"/>
                            <w:lsdException w:name="List Table 2 Accent 2" w:uiPriority="47"/>
                            <w:lsdException w:name="List Table 3 Accent 2" w:uiPriority="48"/>
                            <w:lsdException w:name="List Table 4 Accent 2" w:uiPriority="49"/>
                            <w:lsdException w:name="List Table 5 Dark Accent 2" w:uiPriority="50"/>
                            <w:lsdException w:name="List Table 6 Colorful Accent 2" w:uiPriority="51"/>
                            <w:lsdException w:name="List Table 7 Colorful Accent 2" w:uiPriority="52"/>
                            <w:lsdException w:name="List Table 1 Light Accent 3" w:uiPriority="46"/>
                            <w:lsdException w:name="List Table 2 Accent 3" w:uiPriority="47"/>
                            <w:lsdException w:name="List Table 3 Accent 3" w:uiPriority="48"/>
                            <w:lsdException w:name="List Table 4 Accent 3" w:uiPriority="49"/>
                            <w:lsdException w:name="List Table 5 Dark Accent 3" w:uiPriority="50"/>
                            <w:lsdException w:name="List Table 6 Colorful Accent 3" w:uiPriority="51"/>
                            <w:lsdException w:name="List Table 7 Colorful Accent 3" w:uiPriority="52"/>
                            <w:lsdException w:name="List Table 1 Light Accent 4" w:uiPriority="46"/>
                            <w:lsdException w:name="List Table 2 Accent 4" w:uiPriority="47"/>
                            <w:lsdException w:name="List Table 3 Accent 4" w:uiPriority="48"/>
                            <w:lsdException w:name="List Table 4 Accent 4" w:uiPriority="49"/>
                            <w:lsdException w:name="List Table 5 Dark Accent 4" w:uiPriority="50"/>
                            <w:lsdException w:name="List Table 6 Colorful Accent 4" w:uiPriority="51"/>
                            <w:lsdException w:name="List Table 7 Colorful Accent 4" w:uiPriority="52"/>
                            <w:lsdException w:name="List Table 1 Light Accent 5" w:uiPriority="46"/>
                            <w:lsdException w:name="List Table 2 Accent 5" w:uiPriority="47"/>
                            <w:lsdException w:name="List Table 3 Accent 5" w:uiPriority="48"/>
                            <w:lsdException w:name="List Table 4 Accent 5" w:uiPriority="49"/>
                            <w:lsdException w:name="List Table 5 Dark Accent 5" w:uiPriority="50"/>
                            <w:lsdException w:name="List Table 6 Colorful Accent 5" w:uiPriority="51"/>
                            <w:lsdException w:name="List Table 7 Colorful Accent 5" w:uiPriority="52"/>
                            <w:lsdException w:name="List Table 1 Light Accent 6" w:uiPriority="46"/>
                            <w:lsdException w:name="List Table 2 Accent 6" w:uiPriority="47"/>
                            <w:lsdException w:name="List Table 3 Accent 6" w:uiPriority="48"/>
                            <w:lsdException w:name="List Table 4 Accent 6" w:uiPriority="49"/>
                            <w:lsdException w:name="List Table 5 Dark Accent 6" w:uiPriority="50"/>
                            <w:lsdException w:name="List Table 6 Colorful Accent 6" w:uiPriority="51"/>
                            <w:lsdException w:name="List Table 7 Colorful Accent 6" w:uiPriority="52"/>
                            <w:lsdException w:name="Mention" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Smart Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Hashtag" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Unresolved Mention" w:semiHidden="1" w:unhideWhenUsed="1"/>
                            <w:lsdException w:name="Smart Link" w:semiHidden="1" w:unhideWhenUsed="1"/>
                        </w:latentStyles>
                        <w:style w:type="paragraph" w:default="1" w:styleId="a">
                            <w:name w:val="Normal"/>
                            <w:qFormat/>
                            <w:pPr>
                                <w:widowControl w:val="0"/>
                                <w:jc w:val="both"/>
                            </w:pPr>
                            <w:rPr>
                                <w:kern w:val="2"/>
                                <w:sz w:val="21"/>
                                <w:szCs w:val="22"/>
                            </w:rPr>
                        </w:style>
                        <w:style w:type="character" w:default="1" w:styleId="a0">
                            <w:name w:val="Default Paragraph Font"/>
                            <w:uiPriority w:val="1"/>
                            <w:semiHidden/>
                            <w:unhideWhenUsed/>
                        </w:style>
                        <w:style w:type="table" w:default="1" w:styleId="a1">
                            <w:name w:val="Normal Table"/>
                            <w:uiPriority w:val="99"/>
                            <w:semiHidden/>
                            <w:unhideWhenUsed/>
                            <w:tblPr>
                                <w:tblInd w:w="0" w:type="dxa"/>
                                <w:tblCellMar>
                                    <w:top w:w="0" w:type="dxa"/>
                                    <w:left w:w="108" w:type="dxa"/>
                                    <w:bottom w:w="0" w:type="dxa"/>
                                    <w:right w:w="108" w:type="dxa"/>
                                </w:tblCellMar>
                            </w:tblPr>
                        </w:style>
                        <w:style w:type="numbering" w:default="1" w:styleId="a2">
                            <w:name w:val="No List"/>
                            <w:uiPriority w:val="99"/>
                            <w:semiHidden/>
                            <w:unhideWhenUsed/>
                        </w:style>
                    </w:styles>
                    """;
            zipOutputStream.putNextEntry(new ZipEntry("word/styles.xml"));
            zipOutputStream.write(data7.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建word/styles.xml");
            //生成word/webSettings.xml
            String data8 = """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <w:webSettings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                   xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                                   xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                                   xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                                   xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
                                   xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
                                   xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
                                   xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
                                   xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
                                   mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh"/>
                    """;
            zipOutputStream.putNextEntry(new ZipEntry("word/webSettings.xml"));
            zipOutputStream.write(data8.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建word/webSettings.xml");
            //生成[Content_Types].xml
            String data9 = """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                        <Default Extension="jpeg" ContentType="image/jpeg"/>
                        <Default Extension="png" ContentType="image/png"/>
                        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                        <Default Extension="xml" ContentType="application/xml"/>
                        <Override PartName="/word/document.xml"
                                  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                        <Override PartName="/word/styles.xml"
                                  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                        <Override PartName="/word/settings.xml"
                                  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
                        <Override PartName="/word/webSettings.xml"
                                  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
                        <Override PartName="/word/fontTable.xml"
                                  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
                        <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                        <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
                        <Override PartName="/docProps/app.xml"
                                  ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
                        <Override PartName="/docProps/custom.xml"
                                  ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
                    </Types>
                    """;
            zipOutputStream.putNextEntry(new ZipEntry("[Content_Types].xml"));
            zipOutputStream.write(data9.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建[Content_Types].xml");
            //生成_rels/.rels文件
            String data10= """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                        <Relationship Id="rId3"
                                      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                                      Target="docProps/app.xml"/>
                        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                                      Target="docProps/core.xml"/>
                        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                                      Target="word/document.xml"/>
                        <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
                                      Target="docProps/custom.xml"/>
                    </Relationships>
                    """;
            zipOutputStream.putNextEntry(new ZipEntry("_rels/.rels"));
            zipOutputStream.write(data10.getBytes());
            zipOutputStream.closeEntry();
            System.out.println("[+] 成功创建_rels/.rels文件");
        }
        catch (IOException e){
            System.out.println("[-] 相关文件创建失败，docx文件可能损坏");
            throw e;
        }
    }

    public DocxGenerater(String dest, Reader mdReader) throws FileNotFoundException {
        System.out.println(dest);
        if (dest.matches("(docx)$")) {
            throw new FileNotFoundException("目标不是docx文件");
        }
        docx = new File(dest);
        source = mdReader;
    }

    public void getDocx() throws IOException{
        FileOutputStream outputStream = new FileOutputStream(docx);
        ZipOutputStream zipOutputStream = new ZipOutputStream(outputStream);
        try {
            // 创建.docx文件
            mainContentGen(zipOutputStream);
            relatedGen(zipOutputStream);
            // 关闭输出流
            zipOutputStream.close();
            outputStream.close();
            System.out.println("[+] 成功创建docx文件！");
        } catch (IOException e) {
            zipOutputStream.flush();
            zipOutputStream.close();
            outputStream.close();
            System.out.println("[-] 创建docx文件失败，请检查相关资源……");
            throw e;
        }
    }


    private void createParagraph(ZipOutputStream zipOutputStream, String par, int idx) throws IOException {
        String p = String.format("""
                <w:p w14:paraId="%s" w14:textId="%s" w:rsidR="00F34C9C" w:rsidRDefault="00F34C9C">
                    <w:r>
                        <w:t>%s</w:t>
                    </w:r>
                </w:p>
                """, createPID(idx), createPID(idx), par);
        zipOutputStream.write(p.getBytes());
    }

    private void createPicture(ZipOutputStream zipOutputStream, String imgPath, int idx, ArrayList<String> imgRefs) throws IOException {
        File insertImg = new File(imgPath);
        FileInputStream fin = new FileInputStream(insertImg);
        BufferedImage bImage = ImageIO.read(fin);
        int widthPX = bImage.getWidth();
        int heightPx = bImage.getHeight();
        BigDecimal widthDots = BigDecimal.valueOf(widthPX).multiply(BigDecimal.valueOf(96)).multiply(BigDecimal.valueOf(41));
        BigDecimal heightDots = BigDecimal.valueOf(heightPx).multiply(BigDecimal.valueOf(96)).multiply(BigDecimal.valueOf(41));
        String picId = createPicID(idx);
        imgRefs.add(picId);
        String p = String.format("""
                <w:p>
                    <w:r>
                        <w:rPr>
                            <w:rFonts w:hint="eastAsia"/>
                            <w:noProof/>
                        </w:rPr>
                        <w:drawing>
                            <wp:inline distT="0" distB="0" distL="114300" distR="114300" wp14:anchorId="4B525B94"
                                       wp14:editId="319021F5">
                                <wp:extent cx="%s" cy="%s"/>
                                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                                <wp:docPr id="1" name="图片 1" descr="12"/>
                                <wp:cNvGraphicFramePr>
                                    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                                         noChangeAspect="1"/>
                                </wp:cNvGraphicFramePr>
                                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                            <pic:nvPicPr>
                                                <pic:cNvPr id="1" name="图片 1" descr="12"/>
                                                <pic:cNvPicPr>
                                                    <a:picLocks noChangeAspect="1"/>
                                                </pic:cNvPicPr>
                                            </pic:nvPicPr>
                                            <pic:blipFill>
                                                <a:blip r:embed="%s"/>
                                                <a:stretch>
                                                    <a:fillRect/>
                                                </a:stretch>
                                            </pic:blipFill>
                                            <pic:spPr>
                                                <a:xfrm>
                                                    <a:off x="0" y="0"/>
                                                    <a:ext cx="%s" cy="%s"/>
                                                </a:xfrm>
                                                <a:prstGeom prst="rect">
                                                    <a:avLst/>
                                                </a:prstGeom>
                                            </pic:spPr>
                                        </pic:pic>
                                    </a:graphicData>
                                </a:graphic>
                            </wp:inline>
                        </w:drawing>
                    </w:r>
                </w:p>
                """, widthDots.toString(), heightDots.toString(), picId, widthDots.toString(), heightDots.toString());
        zipOutputStream.write(p.getBytes());
    }

    private void createMedia(ZipOutputStream zipOutputStream, ArrayList<String> imgAlas) throws IOException {
        var imgUrls = source.getImgURLs();
        for (int i = 1; i <= imgUrls.size(); i++) {
            var fileName="media/image" + i + getFileExt(imgUrls.get(i - 1));
            var entryName = "word/"+fileName;
            imgAlas.add(fileName);
            zipOutputStream.putNextEntry(new ZipEntry(entryName));
            FileInputStream fin = new FileInputStream(imgUrls.get(i-1));
            byte[] data = new byte[fin.available()];
            fin.read(data);
            zipOutputStream.write(data);
            zipOutputStream.closeEntry();
        }
    }

    private void createRefs(ZipOutputStream zipOutputStream, ArrayList<String> imgRefs, ArrayList<String> imgAlas) throws IOException {
        String header = """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
                                  Target="webSettings.xml"/>
                    <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                                  Target="theme/theme1.xml"/>
                    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
                                  Target="settings.xml"/>
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                                  Target="styles.xml"/>
                    <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
                                  Target="fontTable.xml"/>
                """;
        StringBuffer sdata = new StringBuffer(header);
        for (int i = 0; i < imgRefs.size(); i++) {
            String relation = String.format("""
                    <Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                                      Target="%s"/>
                    """, imgRefs.get(i), imgAlas.get(i));
            sdata.append(relation);
        }
        sdata.append("</Relationships>");
        zipOutputStream.putNextEntry(new ZipEntry("word/_rels/document.xml.rels"));
        zipOutputStream.write(sdata.toString().getBytes());
        zipOutputStream.closeEntry();
    }

    private String createPID(int idx) {
        String hex = Integer.toHexString(idx);
        StringBuffer ret = new StringBuffer(8);
        int length = 8 - hex.length();
        while (length > 0) {
            ret.append('0');
            length--;
        }
        ret.append(hex);
        return ret.toString();
    }

    private String createPicID(int idx) {
        String hex = Integer.toHexString(idx+57343);
        hex=hex.substring(0,4);
        StringBuffer ret = new StringBuffer(4);
        ret.append(hex);
        return ret.toString();
    }

    private String getFileExt(String path) {
        return path.endsWith(".png") ? ".png" : ".jpeg";
    }
}


