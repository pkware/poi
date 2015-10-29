/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.apache.poi.xslf.usermodel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertSame;
import static org.junit.Assert.assertTrue;

import java.awt.Color;
import java.io.IOException;
import java.util.List;

import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.XSLFTestDataSamples;
import org.junit.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableCell;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGraphicalObjectFrame;

/**
 * @author Yegor Kozlov
 */
public class TestXSLFTable {
    @Test
    public void testRead() throws IOException {
        XMLSlideShow  ppt = XSLFTestDataSamples.openSampleDocument("shapes.pptx");

        XSLFSlide slide = ppt.getSlides().get(3);
        List<XSLFShape> shapes = slide.getShapes();
        assertEquals(1, shapes.size());
        assertTrue(shapes.get(0) instanceof XSLFTable);
        XSLFTable tbl = (XSLFTable)shapes.get(0);
        assertEquals(3, tbl.getNumberOfColumns());
        assertEquals(6, tbl.getNumberOfRows());
        assertNotNull(tbl.getCTTable());

        List<XSLFTableRow> rows = tbl.getRows();
        assertEquals(6, rows.size());

        assertEquals(90.0, tbl.getColumnWidth(0), 0);
        assertEquals(240.0, tbl.getColumnWidth(1), 0);
        assertEquals(150.0, tbl.getColumnWidth(2), 0);

        for(XSLFTableRow row : tbl){
            // all rows have the same height
            assertEquals(29.2, row.getHeight(), 0);
        }

        XSLFTableRow row0 = rows.get(0);
        List<XSLFTableCell> cells0 = row0.getCells();
        assertEquals(3, cells0.size());
        assertEquals("header1", cells0.get(0).getText());
        assertEquals("header2", cells0.get(1).getText());
        assertEquals("header3", cells0.get(2).getText());

        XSLFTableRow row1 = rows.get(1);
        List<XSLFTableCell> cells1 = row1.getCells();
        assertEquals(3, cells1.size());
        assertEquals("A1", cells1.get(0).getText());
        assertEquals("B1", cells1.get(1).getText());
        assertEquals("C1", cells1.get(2).getText());
        
        ppt.close();
    }

    @Test
    public void testCreate() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();

        XSLFTable tbl = slide.createTable();
        assertNotNull(tbl.getCTTable());
        assertNotNull(tbl.getCTTable().getTblGrid());
        assertNotNull(tbl.getCTTable().getTblPr());
        assertTrue(tbl.getXmlObject() instanceof CTGraphicalObjectFrame);
        assertEquals("Table 1", tbl.getShapeName());
        assertEquals(2, tbl.getShapeId());
        assertEquals(0, tbl.getRows().size());
        assertEquals(0, tbl.getCTTable().sizeOfTrArray());
        assertEquals(0, tbl.getCTTable().getTblGrid().sizeOfGridColArray());

        assertEquals(0, tbl.getNumberOfColumns());
        assertEquals(0, tbl.getNumberOfRows());

        XSLFTableRow row0 = tbl.addRow();
        assertNotNull(row0.getXmlObject());
        assertEquals(1, tbl.getNumberOfRows());
        assertSame(row0, tbl.getRows().get(0));
        assertEquals(20.0, row0.getHeight(), 0);
        row0.setHeight(30.0);
        assertEquals(30.0, row0.getHeight(), 0);

        assertEquals(0, row0.getCells().size());
        XSLFTableCell cell0 = row0.addCell();
        assertNotNull(cell0.getXmlObject());
        // by default table cell has no borders
        CTTableCell tc = (CTTableCell)cell0.getXmlObject();
        assertTrue(tc.getTcPr().getLnB().isSetNoFill());
        assertTrue(tc.getTcPr().getLnT().isSetNoFill());
        assertTrue(tc.getTcPr().getLnL().isSetNoFill());
        assertTrue(tc.getTcPr().getLnR().isSetNoFill());

        assertSame(cell0, row0.getCells().get(0));
        assertEquals(1, tbl.getNumberOfColumns());
        assertEquals(100.0, tbl.getColumnWidth(0), 0);
        cell0.addNewTextParagraph().addNewTextRun().setText("POI");
        assertEquals("POI", cell0.getText());

        XSLFTableCell cell1 = row0.addCell();
        assertSame(cell1, row0.getCells().get(1));
        assertEquals(2, tbl.getNumberOfColumns());
        assertEquals(100.0, tbl.getColumnWidth(1), 0);
        cell1.addNewTextParagraph().addNewTextRun().setText("Apache");
        assertEquals("Apache", cell1.getText());

        for (BorderEdge edge : BorderEdge.values()) {
            assertNull(cell1.getBorderWidth(edge));
            cell1.setBorderWidth(edge, 2.0);
            assertEquals(2.0, cell1.getBorderWidth(edge), 0);
            assertNull(cell1.getBorderColor(edge));
            cell1.setBorderColor(edge, Color.yellow);
            assertEquals(Color.yellow, cell1.getBorderColor(edge));
        }

        assertEquals(VerticalAlignment.TOP, cell1.getVerticalAlignment());
        cell1.setVerticalAlignment(VerticalAlignment.MIDDLE);
        assertEquals(VerticalAlignment.MIDDLE, cell1.getVerticalAlignment());
        cell1.setVerticalAlignment(null);
        assertEquals(VerticalAlignment.TOP, cell1.getVerticalAlignment());
        
        ppt.close();
    }
}