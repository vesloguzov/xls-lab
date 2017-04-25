# -*- coding: UTF-8 -*-
import sys
import random
import datetime

from openpyxl import Workbook, load_workbook
from openpyxl import Workbook, load_workbook
from openpyxl.styles import *
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.reader import reader
from openpyxl.chart.layout import Layout, ManualLayout

reload(sys)
sys.setdefaultencoding('utf8')


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
                     <a:rPr lang="ru-RU" />
                     <a:t>График функции</a:t>
                  </a:r>
                  <a:r>
                     <a:rPr lang="en-US" />
                     <a:t>y=sin(x)</a:t>
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
         <c:layout>
            <c:manualLayout>
               <c:layoutTarget val="inner" />
               <c:xMode val="edge" />
               <c:yMode val="edge" />
               <c:x val="3.350972282428185E-2" />
               <c:y val="0.2164883478199317" />
               <c:w val="0.91298064458807493" />
               <c:h val="0.72643963868185513" />
            </c:manualLayout>
         </c:layout>
         <c:scatterChart>
            <c:scatterStyle val="smoothMarker" />
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
                     <c:f>'Лист 1'!$A$4:$A$28</c:f>
                     <c:numCache>
                        <c:formatCode>General</c:formatCode>
                        <c:ptCount val="25" />
                        <c:pt idx="0">
                           <c:v>-6</c:v>
                        </c:pt>
                        <c:pt idx="1">
                           <c:v>-5.5</c:v>
                        </c:pt>
                        <c:pt idx="2">
                           <c:v>-5</c:v>
                        </c:pt>
                        <c:pt idx="3">
                           <c:v>-4.5</c:v>
                        </c:pt>
                        <c:pt idx="4">
                           <c:v>-4</c:v>
                        </c:pt>
                        <c:pt idx="5">
                           <c:v>-3.5</c:v>
                        </c:pt>
                        <c:pt idx="6">
                           <c:v>-3</c:v>
                        </c:pt>
                        <c:pt idx="7">
                           <c:v>-2.5</c:v>
                        </c:pt>
                        <c:pt idx="8">
                           <c:v>-2</c:v>
                        </c:pt>
                        <c:pt idx="9">
                           <c:v>-1.5</c:v>
                        </c:pt>
                        <c:pt idx="10">
                           <c:v>-1</c:v>
                        </c:pt>
                        <c:pt idx="11">
                           <c:v>-0.5</c:v>
                        </c:pt>
                        <c:pt idx="12">
                           <c:v>0</c:v>
                        </c:pt>
                        <c:pt idx="13">
                           <c:v>0.5</c:v>
                        </c:pt>
                        <c:pt idx="14">
                           <c:v>1</c:v>
                        </c:pt>
                        <c:pt idx="15">
                           <c:v>1.5</c:v>
                        </c:pt>
                        <c:pt idx="16">
                           <c:v>2</c:v>
                        </c:pt>
                        <c:pt idx="17">
                           <c:v>2.5</c:v>
                        </c:pt>
                        <c:pt idx="18">
                           <c:v>3</c:v>
                        </c:pt>
                        <c:pt idx="19">
                           <c:v>3.5</c:v>
                        </c:pt>
                        <c:pt idx="20">
                           <c:v>4</c:v>
                        </c:pt>
                        <c:pt idx="21">
                           <c:v>4.5</c:v>
                        </c:pt>
                        <c:pt idx="22">
                           <c:v>5</c:v>
                        </c:pt>
                        <c:pt idx="23">
                           <c:v>5.5</c:v>
                        </c:pt>
                        <c:pt idx="24">
                           <c:v>6</c:v>
                        </c:pt>
                     </c:numCache>
                  </c:numRef>
               </c:xVal>
               <c:yVal>
                  <c:numRef>
                     <c:f>'Лист 1'!$B$4:$B$28</c:f>
                     <c:numCache>
                        <c:formatCode>General</c:formatCode>
                        <c:ptCount val="25" />
                        <c:pt idx="0">
                           <c:v>0.27941549819892586</c:v>
                        </c:pt>
                        <c:pt idx="1">
                           <c:v>0.70554032557039192</c:v>
                        </c:pt>
                        <c:pt idx="2">
                           <c:v>0.95892427466313845</c:v>
                        </c:pt>
                        <c:pt idx="3">
                           <c:v>0.97753011766509701</c:v>
                        </c:pt>
                        <c:pt idx="4">
                           <c:v>0.7568024953079282</c:v>
                        </c:pt>
                        <c:pt idx="5">
                           <c:v>0.35078322768961984</c:v>
                        </c:pt>
                        <c:pt idx="6">
                           <c:v>-0.14112000805986721</c:v>
                        </c:pt>
                        <c:pt idx="7">
                           <c:v>-0.59847214410395655</c:v>
                        </c:pt>
                        <c:pt idx="8">
                           <c:v>-0.90929742682568171</c:v>
                        </c:pt>
                        <c:pt idx="9">
                           <c:v>-0.99749498660405445</c:v>
                        </c:pt>
                        <c:pt idx="10">
                           <c:v>-0.8414709848078965</c:v>
                        </c:pt>
                        <c:pt idx="11">
                           <c:v>-0.47942553860420301</c:v>
                        </c:pt>
                        <c:pt idx="12">
                           <c:v>0</c:v>
                        </c:pt>
                        <c:pt idx="13">
                           <c:v>0.47942553860420301</c:v>
                        </c:pt>
                        <c:pt idx="14">
                           <c:v>0.8414709848078965</c:v>
                        </c:pt>
                        <c:pt idx="15">
                           <c:v>0.99749498660405445</c:v>
                        </c:pt>
                        <c:pt idx="16">
                           <c:v>0.90929742682568171</c:v>
                        </c:pt>
                        <c:pt idx="17">
                           <c:v>0.59847214410395655</c:v>
                        </c:pt>
                        <c:pt idx="18">
                           <c:v>0.14112000805986721</c:v>
                        </c:pt>
                        <c:pt idx="19">
                           <c:v>-0.35078322768961984</c:v>
                        </c:pt>
                        <c:pt idx="20">
                           <c:v>-0.7568024953079282</c:v>
                        </c:pt>
                        <c:pt idx="21">
                           <c:v>-0.97753011766509701</c:v>
                        </c:pt>
                        <c:pt idx="22">
                           <c:v>-0.95892427466313845</c:v>
                        </c:pt>
                        <c:pt idx="23">
                           <c:v>-0.70554032557039192</c:v>
                        </c:pt>
                        <c:pt idx="24">
                           <c:v>-0.27941549819892586</c:v>
                        </c:pt>
                     </c:numCache>
                  </c:numRef>
               </c:yVal>
               <c:smooth val="1" />
            </c:ser>
            <c:dLbls>
               <c:showLegendKey val="0" />
               <c:showVal val="0" />
               <c:showCatName val="0" />
               <c:showSerName val="0" />
               <c:showPercent val="0" />
               <c:showBubbleSize val="0" />
            </c:dLbls>
            <c:axId val="19787119" />
            <c:axId val="19787951" />
         </c:scatterChart>
         <c:valAx>
            <c:axId val="19787119" />
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
            <c:title>
               <c:tx>
                  <c:rich>
                     <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1" />
                     <a:lstStyle />
                     <a:p>
                        <a:pPr>
                           <a:defRPr sz="2000" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
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
                           <a:rPr lang="ru-RU" sz="2000" />
                           <a:t>осьх</a:t>
                        </a:r>
                        <a:r>
                           <a:rPr lang="en-US" sz="2000" />
                           <a:t>x</a:t>
                        </a:r>
                        <a:endParaRPr lang="ru-RU" sz="2000" />
                     </a:p>
                  </c:rich>
               </c:tx>
               <c:layout>
                  <c:manualLayout>
                     <c:xMode val="edge" />
                     <c:yMode val="edge" />
                     <c:x val="0.95882779829890052" />
                     <c:y val="0.51937322334839675" />
                  </c:manualLayout>
               </c:layout>
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
                        <a:defRPr sz="2000" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
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
            <c:crossAx val="19787951" />
            <c:crosses val="autoZero" />
            <c:crossBetween val="midCat" />
         </c:valAx>
         <c:valAx>
            <c:axId val="19787951" />
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
            <c:title>
               <c:tx>
                  <c:rich>
                     <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1" />
                     <a:lstStyle />
                     <a:p>
                        <a:pPr>
                           <a:defRPr sz="2000" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
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
                           <a:rPr lang="ru-RU" sz="2000" />
                           <a:t>осьу</a:t>
                        </a:r>
                        <a:r>
                           <a:rPr lang="en-US" sz="2000" />
                           <a:t>y</a:t>
                        </a:r>
                        <a:endParaRPr lang="ru-RU" sz="2000" />
                     </a:p>
                  </c:rich>
               </c:tx>
               <c:layout>
                  <c:manualLayout>
                     <c:xMode val="edge" />
                     <c:yMode val="edge" />
                     <c:x val="0.4694214061436574" />
                     <c:y val="0.10315494869256478" />
                  </c:manualLayout>
               </c:layout>
               <c:overlay val="0" />
               <c:spPr>
                  <a:noFill />
                  <a:ln>
                     <a:noFill />
                  </a:ln>
                  <a:effectLst />
               </c:spPr>
               <c:txPr>
                  <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1" />
                  <a:lstStyle />
                  <a:p>
                     <a:pPr>
                        <a:defRPr sz="2000" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
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
            <c:crossAx val="19787119" />
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

# for p in l.title.tx.rich.p:
#     print p.r


# for p in l.y_axis.title.tx:
for p in l.y_axis.title.tx.rich.p:
    print p

print 'KEK'
# print l.x_axis

for s in l.ser:
    # print s.xVal.numRef.numCache # Значения числовые

    print s.xVal.numRef.f # Диапазон данных
    print s.yVal.numRef.f