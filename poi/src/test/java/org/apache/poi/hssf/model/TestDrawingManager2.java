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

package org.apache.poi.hssf.model;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.apache.poi.ddf.EscherBSERecord;
import org.apache.poi.ddf.EscherDgRecord;
import org.apache.poi.ddf.EscherDggRecord;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

final class TestDrawingManager2 {
    private DrawingManager2 drawingManager2;
    private EscherDggRecord dgg;

    @BeforeEach
    void setUp() {
        drawingManager2 = new HSSFWorkbook().getWorkbook().getDrawingManager();
        dgg = drawingManager2.getDgg();
    }

    @Test
    void testCreateDgRecord() {
        EscherDgRecord dgRecord1 = drawingManager2.createDgRecord();
        assertEquals( 1, dgRecord1.getDrawingGroupId() );
        assertEquals( -1, dgRecord1.getLastMSOSPID() );

        EscherDgRecord dgRecord2 = drawingManager2.createDgRecord();
        assertEquals( 2, dgRecord2.getDrawingGroupId() );
        assertEquals( -1, dgRecord2.getLastMSOSPID() );

        assertEquals( 2, dgg.getDrawingsSaved( ) );
        assertEquals( 2, dgg.getFileIdClusters().length );
        assertEquals( 3, dgg.getNumIdClusters() );
        assertEquals( 0, dgg.getNumShapesSaved() );
    }

    @Test
    void testCreateDgRecordOld() {
        // converted from TestDrawingManager(1)

        EscherDgRecord dgRecord = drawingManager2.createDgRecord();
        assertEquals( -1, dgRecord.getLastMSOSPID() );
        assertEquals( 0, dgRecord.getNumShapes() );
        assertEquals( 1, drawingManager2.getDgg().getDrawingsSaved() );
        assertEquals( 1, drawingManager2.getDgg().getFileIdClusters().length );
        assertEquals( 1, drawingManager2.getDgg().getFileIdClusters()[0].getDrawingGroupId() );
        assertEquals( 0, drawingManager2.getDgg().getFileIdClusters()[0].getNumShapeIdsUsed() );
    }

    @Test
    void testAllocateShapeId() {
        EscherDgRecord dgRecord1 = drawingManager2.createDgRecord();
        assertEquals( 1, dgg.getDrawingsSaved() );
        EscherDgRecord dgRecord2 = drawingManager2.createDgRecord();
        assertEquals( 2, dgg.getDrawingsSaved() );

        assertEquals( 1024, drawingManager2.allocateShapeId( dgRecord1 ) );
        assertEquals( 1024, dgRecord1.getLastMSOSPID() );
        assertEquals( 1025, dgg.getShapeIdMax() );
        assertEquals( 1, dgg.getFileIdClusters()[0].getDrawingGroupId() );
        assertEquals( 1, dgg.getFileIdClusters()[0].getNumShapeIdsUsed() );
        assertEquals( 1, dgRecord1.getNumShapes() );
        assertEquals( 1025, drawingManager2.allocateShapeId( dgRecord1 ) );
        assertEquals( 1025, dgRecord1.getLastMSOSPID() );
        assertEquals( 1026, dgg.getShapeIdMax() );
        assertEquals( 1026, drawingManager2.allocateShapeId( dgRecord1 ) );
        assertEquals( 1026, dgRecord1.getLastMSOSPID() );
        assertEquals( 1027, dgg.getShapeIdMax() );
        assertEquals( 2048, drawingManager2.allocateShapeId( dgRecord2 ) );
        assertEquals( 2048, dgRecord2.getLastMSOSPID() );
        assertEquals( 2049, dgg.getShapeIdMax() );

        for (int i = 0; i < 1021; i++)
        {
            drawingManager2.allocateShapeId( dgRecord1 );
            assertEquals( 2049, dgg.getShapeIdMax() );
        }
        assertEquals( 3072, drawingManager2.allocateShapeId( dgRecord1 ) );
        assertEquals( 3073, dgg.getShapeIdMax() );

        assertEquals( 2, dgg.getDrawingsSaved() );
        assertEquals( 4, dgg.getNumIdClusters() );
        assertEquals( 1026, dgg.getNumShapesSaved() );
    }

    @Test
    void testFindNewDrawingGroupId() {
        // converted from TestDrawingManager(1)
        dgg.setDrawingsSaved( 1 );
        dgg.setFileIdClusters( new EscherDggRecord.FileIdCluster[]{
            new EscherDggRecord.FileIdCluster( 2, 10 )} );
        assertEquals( 1, drawingManager2.findNewDrawingGroupId() );
        dgg.setFileIdClusters( new EscherDggRecord.FileIdCluster[]{
            new EscherDggRecord.FileIdCluster( 1, 10 ),
            new EscherDggRecord.FileIdCluster( 2, 10 )} );
        assertEquals( 3, drawingManager2.findNewDrawingGroupId() );
    }

    /**
     * Verify that the {@link DrawingManager2#getPictureCount()} function returns the correct count for various states.
     */
    @Test
    void getPictureCount() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        DrawingManager2 drawingManager = workbook.getInternalWorkbook().getDrawingManager();

        // No pictures
        assertEquals(0, drawingManager.getPictureCount());

        // 1 Picture
        // Picture content and type don't matter
        workbook.addPicture(new byte[10], HSSFWorkbook.PICTURE_TYPE_JPEG);
        assertEquals(1, drawingManager.getPictureCount());

        // 2 Pictures
        // Picture content and type don't matter
        workbook.addPicture(new byte[50], HSSFWorkbook.PICTURE_TYPE_PNG);
        assertEquals(2, drawingManager.getPictureCount());

        // From serialized workbook
        ByteArrayOutputStream inMemory = new ByteArrayOutputStream();
        workbook.write(inMemory);

        workbook = new HSSFWorkbook(new ByteArrayInputStream(inMemory.toByteArray()));
        drawingManager = workbook.getInternalWorkbook().getDrawingManager();
        assertEquals(2, drawingManager.getPictureCount());
    }

    /**
     * Verify that the {@link DrawingManager2#getAllPictures()} ()} function returns all pictures in the right order.
     */
    @Test
    void getAllPictures() throws IOException {
        ByteArrayOutputStream inMemory = new ByteArrayOutputStream();
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            DrawingManager2 drawingManager = workbook.getInternalWorkbook().getDrawingManager();

            // No pictures
            assertTrue(drawingManager.getAllPictures().isEmpty());

            // 1 Picture
            // Picture content and type don't matter
            workbook.addPicture(new byte[10], HSSFWorkbook.PICTURE_TYPE_JPEG);
            assertEquals(1, drawingManager.getAllPictures().size());

            // 2 Pictures
            // Picture content and type don't matter
            workbook.addPicture(new byte[50], HSSFWorkbook.PICTURE_TYPE_PNG);
            List<HSSFPictureData> pictures = drawingManager.getAllPictures();
            assertEquals(2, pictures.size());
            assertEquals(10, pictures.get(0).getData().length);
            assertEquals(50, pictures.get(1).getData().length);

            workbook.write(inMemory);
        }

        // From serialized workbook
        try (HSSFWorkbook workbook = new HSSFWorkbook(new ByteArrayInputStream(inMemory.toByteArray()))) {
            DrawingManager2 drawingManager = workbook.getInternalWorkbook().getDrawingManager();
            List<HSSFPictureData> pictures = drawingManager.getAllPictures();
            assertEquals(2, pictures.size());
            assertEquals(10, pictures.get(0).getData().length);
            assertEquals(50, pictures.get(1).getData().length);
        }
    }

    /**
     * Verify that new BSE records for metafiles have the right tag.
     */
    @ParameterizedTest
    @ValueSource(ints = {Workbook.PICTURE_TYPE_EMF, Workbook.PICTURE_TYPE_WMF})
    void allocatePicture_metafiles(int format) throws IOException {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            DrawingManager2 drawingManager = workbook.getInternalWorkbook().getDrawingManager();
            HSSFPictureData picture = drawingManager.allocatePicture(format);
            assertEquals(format, picture.getFormat());
            EscherBSERecord record = drawingManager.getBSERecord(1);
            assertEquals(0, record.getTag());
        }
    }

    /**
     * Verify that new BSE records for bitmaps have the right tag.
     */
    @ParameterizedTest
    @ValueSource(ints = {
        Workbook.PICTURE_TYPE_JPEG,
        Workbook.PICTURE_TYPE_PICT,
        Workbook.PICTURE_TYPE_PNG,
        Workbook.PICTURE_TYPE_DIB
    })
    void allocatePicture_bitmap(int format) throws IOException {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            DrawingManager2 drawingManager = workbook.getInternalWorkbook().getDrawingManager();
            HSSFPictureData picture = drawingManager.allocatePicture(format);
            assertEquals(format, picture.getFormat());
            EscherBSERecord record = drawingManager.getBSERecord(1);
            assertEquals((short) 0xFF, record.getTag());
        }
    }
}
