# -*- coding: UTF-8 -*-
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.reader import reader
from openpyxl.chart.layout import Layout, ManualLayout
# import xlsxwriter
# import xlwings as xw

reload(sys)
sys.setdefaultencoding('utf8')

wb2 = load_workbook('data/rabota 8.xlsx')
ws = wb2[wb2.get_sheet_names()[0]]

print "LOL"
print ws.sheet_view






















l = reader('''<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
   <c:date1904 val="0" />
   <c:lang val="ru-RU" />
   <c:roundedCorners val="0" />
   <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">
         <c14:style val="102" />
      </mc:Choice>
      <mc:Fallback>
         <c:style val="2" />
      </mc:Fallback>
   </mc:AlternateContent>
   <c:chart>
      <c:title>
         <c:tx>
            <c:rich>
               <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1" />
               <a:lstStyle />
               <a:p>
                  <a:pPr>
                     <a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">
                        <a:solidFill>
                           <a:schemeClr val="tx1">
                              <a:lumMod val="65000" />
                              <a:lumOff val="35000" />
                           </a:schemeClr>
                        </a:solidFill>
                        <a:latin typeface="+mn-lt" />
                        <a:ea typeface="+mn-ea" />
                        <a:cs typeface="+mn-cs" />
                     </a:defRPr>
                  </a:pPr>
                  <a:r>
                     <a:rPr lang="en-US" />
                     <a:t>DIAGRAM NAME</a:t>
                  </a:r>
                  <a:endParaRPr lang="ru-RU" />
               </a:p>
            </c:rich>
         </c:tx>
         <c:layout />
         <c:overlay val="0" />
         <c:spPr>
            <a:noFill />
            <a:ln>
               <a:noFill />
            </a:ln>
            <a:effectLst />
         </c:spPr>
         <c:txPr>
            <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1" />
            <a:lstStyle />
            <a:p>
               <a:pPr>
                  <a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">
                     <a:solidFill>
                        <a:schemeClr val="tx1">
                           <a:lumMod val="65000" />
                           <a:lumOff val="35000" />
                        </a:schemeClr>
                     </a:solidFill>
                     <a:latin typeface="+mn-lt" />
                     <a:ea typeface="+mn-ea" />
                     <a:cs typeface="+mn-cs" />
                  </a:defRPr>
               </a:pPr>
               <a:endParaRPr lang="ru-RU" />
            </a:p>
         </c:txPr>
      </c:title>
      <c:autoTitleDeleted val="0" />
      <c:plotArea>
         <c:layout />
         <c:scatterChart>
            <c:scatterStyle val="lineMarker" />
            <c:varyColors val="0" />
            <c:ser>
               <c:idx val="0" />
               <c:order val="0" />
               <c:spPr>
                  <a:ln w="19050" cap="rnd">
                     <a:solidFill>
                        <a:schemeClr val="accent1" />
                     </a:solidFill>
                     <a:round />
                  </a:ln>
                  <a:effectLst />
               </c:spPr>
               <c:marker>
                  <c:symbol val="circle" />
                  <c:size val="5" />
                  <c:spPr>
                     <a:solidFill>
                        <a:schemeClr val="accent1" />
                     </a:solidFill>
                     <a:ln w="9525">
                        <a:solidFill>
                           <a:schemeClr val="accent1" />
                        </a:solidFill>
                     </a:ln>
                     <a:effectLst />
                  </c:spPr>
               </c:marker>
               <c:xVal>
                  <c:numRef>
                     <c:f>Лист1!$C$3:$C$33</c:f>
                     <c:numCache>
                        <c:formatCode>General</c:formatCode>
                        <c:ptCount val="31" />
                        <c:pt idx="0">
                           <c:v>-3</c:v>
                        </c:pt>
                        <c:pt idx="1">
                           <c:v>-2.8</c:v>
                        </c:pt>
                        <c:pt idx="2">
                           <c:v>-2.6</c:v>
                        </c:pt>
                        <c:pt idx="3">
                           <c:v>-2.4</c:v>
                        </c:pt>
                        <c:pt idx="4">
                           <c:v>-2.2000000000000002</c:v>
                        </c:pt>
                        <c:pt idx="5">
                           <c:v>-2</c:v>
                        </c:pt>
                        <c:pt idx="6">
                           <c:v>-1.8</c:v>
                        </c:pt>
                        <c:pt idx="7">
                           <c:v>-1.6</c:v>
                        </c:pt>
                        <c:pt idx="8">
                           <c:v>-1.4</c:v>
                        </c:pt>
                        <c:pt idx="9">
                           <c:v>-1.2</c:v>
                        </c:pt>
                        <c:pt idx="10">
                           <c:v>-1</c:v>
                        </c:pt>
                        <c:pt idx="11">
                           <c:v>-0.8</c:v>
                        </c:pt>
                        <c:pt idx="12">
                           <c:v>-0.6</c:v>
                        </c:pt>
                        <c:pt idx="13">
                           <c:v>-0.4</c:v>
                        </c:pt>
                        <c:pt idx="14">
                           <c:v>-0.2</c:v>
                        </c:pt>
                        <c:pt idx="15">
                           <c:v>0</c:v>
                        </c:pt>
                        <c:pt idx="16">
                           <c:v>0.2</c:v>
                        </c:pt>
                        <c:pt idx="17">
                           <c:v>0.4</c:v>
                        </c:pt>
                        <c:pt idx="18">
                           <c:v>0.6</c:v>
                        </c:pt>
                        <c:pt idx="19">
                           <c:v>0.8</c:v>
                        </c:pt>
                        <c:pt idx="20">
                           <c:v>1</c:v>
                        </c:pt>
                        <c:pt idx="21">
                           <c:v>1.2</c:v>
                        </c:pt>
                        <c:pt idx="22">
                           <c:v>1.4</c:v>
                        </c:pt>
                        <c:pt idx="23">
                           <c:v>1.6</c:v>
                        </c:pt>
                        <c:pt idx="24">
                           <c:v>1.8</c:v>
                        </c:pt>
                        <c:pt idx="25">
                           <c:v>2</c:v>
                        </c:pt>
                        <c:pt idx="26">
                           <c:v>2.2000000000000002</c:v>
                        </c:pt>
                        <c:pt idx="27">
                           <c:v>2.4</c:v>
                        </c:pt>
                        <c:pt idx="28">
                           <c:v>2.6</c:v>
                        </c:pt>
                        <c:pt idx="29">
                           <c:v>2.80000000000001</c:v>
                        </c:pt>
                        <c:pt idx="30">
                           <c:v>3.0000000000000102</c:v>
                        </c:pt>
                     </c:numCache>
                  </c:numRef>
               </c:xVal>
               <c:yVal>
                  <c:numRef>
                     <c:f>Лист1!$D$3:$D$33</c:f>
                     <c:numCache>
                        <c:formatCode>General</c:formatCode>
                        <c:ptCount val="31" />
                        <c:pt idx="0">
                           <c:v>2</c:v>
                        </c:pt>
                        <c:pt idx="1">
                           <c:v>1.7999999999999998</c:v>
                        </c:pt>
                        <c:pt idx="2">
                           <c:v>1.6</c:v>
                        </c:pt>
                        <c:pt idx="3">
                           <c:v>1.4</c:v>
                        </c:pt>
                        <c:pt idx="4" formatCode="_(&amp;quot;₽&amp;quot;* #,##0.00_);_(&amp;quot;₽&amp;quot;* \(#,##0.00\);_(&amp;quot;₽&amp;quot;* &amp;quot;-&amp;quot;??_);_(@_)">
                           <c:v>1.2000000000000002</c:v>
                        </c:pt>
                        <c:pt idx="5">
                           <c:v>1</c:v>
                        </c:pt>
                        <c:pt idx="6">
                           <c:v>0.8</c:v>
                        </c:pt>
                        <c:pt idx="7">
                           <c:v>0.60000000000000009</c:v>
                        </c:pt>
                        <c:pt idx="8">
                           <c:v>0.39999999999999991</c:v>
                        </c:pt>
                        <c:pt idx="9">
                           <c:v>0.19999999999999996</c:v>
                        </c:pt>
                        <c:pt idx="10">
                           <c:v>0</c:v>
                        </c:pt>
                        <c:pt idx="11">
                           <c:v>0.35999999999999988</c:v>
                        </c:pt>
                        <c:pt idx="12">
                           <c:v>0.64</c:v>
                        </c:pt>
                        <c:pt idx="13">
                           <c:v>0.84</c:v>
                        </c:pt>
                        <c:pt idx="14">
                           <c:v>0.96</c:v>
                        </c:pt>
                        <c:pt idx="15">
                           <c:v>1</c:v>
                        </c:pt>
                        <c:pt idx="16">
                           <c:v>0.96</c:v>
                        </c:pt>
                        <c:pt idx="17">
                           <c:v>0.84</c:v>
                        </c:pt>
                        <c:pt idx="18">
                           <c:v>0.64</c:v>
                        </c:pt>
                        <c:pt idx="19">
                           <c:v>0.35999999999999988</c:v>
                        </c:pt>
                        <c:pt idx="20">
                           <c:v>0</c:v>
                        </c:pt>
                        <c:pt idx="21">
                           <c:v>0.19999999999999996</c:v>
                        </c:pt>
                        <c:pt idx="22">
                           <c:v>0.39999999999999991</c:v>
                        </c:pt>
                        <c:pt idx="23">
                           <c:v>0.60000000000000009</c:v>
                        </c:pt>
                        <c:pt idx="24">
                           <c:v>0.8</c:v>
                        </c:pt>
                        <c:pt idx="25">
                           <c:v>1</c:v>
                        </c:pt>
                        <c:pt idx="26">
                           <c:v>1.2000000000000002</c:v>
                        </c:pt>
                        <c:pt idx="27">
                           <c:v>1.4</c:v>
                        </c:pt>
                        <c:pt idx="28">
                           <c:v>1.6</c:v>
                        </c:pt>
                        <c:pt idx="29">
                           <c:v>1.80000000000001</c:v>
                        </c:pt>
                        <c:pt idx="30">
                           <c:v>2.0000000000000102</c:v>
                        </c:pt>
                     </c:numCache>
                  </c:numRef>
               </c:yVal>
               <c:smooth val="0" />
            </c:ser>
            <c:dLbls>
               <c:showLegendKey val="0" />
               <c:showVal val="0" />
               <c:showCatName val="0" />
               <c:showSerName val="0" />
               <c:showPercent val="0" />
               <c:showBubbleSize val="0" />
            </c:dLbls>
            <c:axId val="1491973487" />
            <c:axId val="1491962671" />
         </c:scatterChart>
         <c:valAx>
            <c:axId val="1491973487" />
            <c:scaling>
               <c:orientation val="minMax" />
            </c:scaling>
            <c:delete val="0" />
            <c:axPos val="b" />
            <c:majorGridlines>
               <c:spPr>
                  <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                     <a:solidFill>
                        <a:schemeClr val="tx1">
                           <a:lumMod val="15000" />
                           <a:lumOff val="85000" />
                        </a:schemeClr>
                     </a:solidFill>
                     <a:round />
                  </a:ln>
                  <a:effectLst />
               </c:spPr>
            </c:majorGridlines>
            <c:numFmt formatCode="General" sourceLinked="1" />
            <c:majorTickMark val="none" />
            <c:minorTickMark val="none" />
            <c:tickLblPos val="nextTo" />
            <c:spPr>
               <a:noFill />
               <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                  <a:solidFill>
                     <a:schemeClr val="tx1">
                        <a:lumMod val="25000" />
                        <a:lumOff val="75000" />
                     </a:schemeClr>
                  </a:solidFill>
                  <a:round />
               </a:ln>
               <a:effectLst />
            </c:spPr>
            <c:txPr>
               <a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1" />
               <a:lstStyle />
               <a:p>
                  <a:pPr>
                     <a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
                        <a:solidFill>
                           <a:schemeClr val="tx1">
                              <a:lumMod val="65000" />
                              <a:lumOff val="35000" />
                           </a:schemeClr>
                        </a:solidFill>
                        <a:latin typeface="+mn-lt" />
                        <a:ea typeface="+mn-ea" />
                        <a:cs typeface="+mn-cs" />
                     </a:defRPr>
                  </a:pPr>
                  <a:endParaRPr lang="ru-RU" />
               </a:p>
            </c:txPr>
            <c:crossAx val="1491962671" />
            <c:crosses val="autoZero" />
            <c:crossBetween val="midCat" />
         </c:valAx>
         <c:valAx>
            <c:axId val="1491962671" />
            <c:scaling>
               <c:orientation val="minMax" />
            </c:scaling>
            <c:delete val="0" />
            <c:axPos val="l" />
            <c:majorGridlines>
               <c:spPr>
                  <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                     <a:solidFill>
                        <a:schemeClr val="tx1">
                           <a:lumMod val="15000" />
                           <a:lumOff val="85000" />
                        </a:schemeClr>
                     </a:solidFill>
                     <a:round />
                  </a:ln>
                  <a:effectLst />
               </c:spPr>
            </c:majorGridlines>
            <c:numFmt formatCode="General" sourceLinked="1" />
            <c:majorTickMark val="none" />
            <c:minorTickMark val="none" />
            <c:tickLblPos val="nextTo" />
            <c:spPr>
               <a:noFill />
               <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                  <a:solidFill>
                     <a:schemeClr val="tx1">
                        <a:lumMod val="25000" />
                        <a:lumOff val="75000" />
                     </a:schemeClr>
                  </a:solidFill>
                  <a:round />
               </a:ln>
               <a:effectLst />
            </c:spPr>
            <c:txPr>
               <a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1" />
               <a:lstStyle />
               <a:p>
                  <a:pPr>
                     <a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
                        <a:solidFill>
                           <a:schemeClr val="tx1">
                              <a:lumMod val="65000" />
                              <a:lumOff val="35000" />
                           </a:schemeClr>
                        </a:solidFill>
                        <a:latin typeface="+mn-lt" />
                        <a:ea typeface="+mn-ea" />
                        <a:cs typeface="+mn-cs" />
                     </a:defRPr>
                  </a:pPr>
                  <a:endParaRPr lang="ru-RU" />
               </a:p>
            </c:txPr>
            <c:crossAx val="1491973487" />
            <c:crosses val="autoZero" />
            <c:crossBetween val="midCat" />
         </c:valAx>
         <c:spPr>
            <a:noFill />
            <a:ln>
               <a:noFill />
            </a:ln>
            <a:effectLst />
         </c:spPr>
      </c:plotArea>
      <c:plotVisOnly val="1" />
      <c:dispBlanksAs val="gap" />
      <c:showDLblsOverMax val="0" />
   </c:chart>
   <c:spPr>
      <a:solidFill>
         <a:schemeClr val="bg1" />
      </a:solidFill>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
         <a:solidFill>
            <a:schemeClr val="tx1">
               <a:lumMod val="15000" />
               <a:lumOff val="85000" />
            </a:schemeClr>
         </a:solidFill>
         <a:round />
      </a:ln>
      <a:effectLst />
   </c:spPr>
   <c:txPr>
      <a:bodyPr />
      <a:lstStyle />
      <a:p>
         <a:pPr>
            <a:defRPr />
         </a:pPr>
         <a:endParaRPr lang="ru-RU" />
      </a:p>
   </c:txPr>
   <c:printSettings>
      <c:headerFooter />
      <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3" />
      <c:pageSetup />
   </c:printSettings>
</c:chartSpace>''')


# Заголовок ScutterChart
# for p in l.title.tx.rich.p:
#     print p.text.t

# for s in l.ser:
#     # print s.xVal.numRef.numCache # Значения числовые
#     print s.xVal.numRef.f # Диапозон данных
#     print s.yVal.numRef.f