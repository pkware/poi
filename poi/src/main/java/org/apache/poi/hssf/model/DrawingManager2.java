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

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ddf.EscherBSERecord;
import org.apache.poi.ddf.EscherBitmapBlip;
import org.apache.poi.ddf.EscherBlipRecord;
import org.apache.poi.ddf.EscherBoolProperty;
import org.apache.poi.ddf.EscherContainerRecord;
import org.apache.poi.ddf.EscherDgRecord;
import org.apache.poi.ddf.EscherDggRecord;
import org.apache.poi.ddf.EscherMetafileBlip;
import org.apache.poi.ddf.EscherOptRecord;
import org.apache.poi.ddf.EscherPropertyTypes;
import org.apache.poi.ddf.EscherRGBProperty;
import org.apache.poi.ddf.EscherRecord;
import org.apache.poi.ddf.EscherRecordTypes;
import org.apache.poi.ddf.EscherSplitMenuColorsRecord;
import org.apache.poi.hssf.record.CountryRecord;
import org.apache.poi.hssf.record.DrawingGroupRecord;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Internal;
import org.apache.poi.util.Removal;

import static org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_DIB;
import static org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_EMF;
import static org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_JPEG;
import static org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_PICT;
import static org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_PNG;
import static org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_WMF;


/**
 * Provides utilities to manage drawing groups.
 */
public class DrawingManager2 {
    private static final Logger LOGGER = LogManager.getLogger(DrawingManager2.class);
    private final EscherDggRecord dgg;

    /**
     * A {@link EscherRecordTypes#BSTORE_CONTAINER}.
     */
    private final EscherContainerRecord bStoreContainer;
    private final List<EscherDgRecord> drawingGroups = new ArrayList<>();

    /**
     * @deprecated This class is intended to be internal to POI. To get instances of it, call
     * {@link InternalWorkbook#getDrawingManager()}.
     */
    @Deprecated
    @Removal(version = "5.4")
    public DrawingManager2( EscherDggRecord dgg ) {
        this(dgg, new EscherContainerRecord());
        LOGGER.atWarn().log("Detached BStore container created = data won't be saved.");
        bStoreContainer.setRecordId(EscherContainerRecord.BSTORE_CONTAINER);
    }

    private DrawingManager2(EscherDggRecord dgg, EscherContainerRecord bStoreContainer) {
        this.dgg = Objects.requireNonNull(dgg);
        this.bStoreContainer = Objects.requireNonNull(bStoreContainer);
    }

    /**
     * Initializes a {@link DrawingManager2} tied to the given workbook.
     * <p>
     * If the workbook does not yet have records for the Drawing Group, BStore, & DGG records, new records will be
     * created and added to the workbook.
     *
     * @param workbook From which to initialize the {@link DrawingManager2}.
     * @return The new {@link DrawingManager2} instance.
     */
    static synchronized DrawingManager2 forWorkbook(InternalWorkbook workbook) {
        DrawingGroupRecord drawingGroup = (DrawingGroupRecord) workbook.findFirstRecordBySid(DrawingGroupRecord.sid);

        if (drawingGroup != null) {
            drawingGroup.decode();
        } else {
            drawingGroup = createDrawingGroupRecord();

            int index = workbook.findFirstRecordLocBySid(CountryRecord.sid);
            workbook.getWorkbookRecordList().add(index + 1, drawingGroup);
        }

        EscherContainerRecord dggContainer = drawingGroup.getEscherContainer();

        EscherDggRecord dggRecord = dggContainer.getChildById(EscherDggRecord.RECORD_ID);
        if (dggRecord == null) {
            dggRecord = createEscherDggRecord();
            dggContainer.addChildRecord(dggRecord);
        }

        EscherContainerRecord bStoreContainer = dggContainer.getChildById(EscherContainerRecord.BSTORE_CONTAINER);
        if (bStoreContainer == null) {
            // Create an empty store and use that. An empty store doesn't cause any problems, but makes the programmatic
            //  model easier to understand.
            bStoreContainer = new EscherContainerRecord();
            bStoreContainer.setRecordId(EscherContainerRecord.BSTORE_CONTAINER);
            bStoreContainer.setOptions((short) (0xF));

            List<EscherRecord> childRecords = dggContainer.getChildRecords();
            childRecords.add(1, bStoreContainer);
            dggContainer.setChildRecords(childRecords);
        }

        return new DrawingManager2(dggRecord, bStoreContainer);
    }

    /**
     * Creates a primary drawing group record.
     */
    private static DrawingGroupRecord createDrawingGroupRecord() {
        EscherContainerRecord dggContainer = new EscherContainerRecord();
        EscherDggRecord dgg = createEscherDggRecord();
        EscherOptRecord opt = new EscherOptRecord();
        EscherSplitMenuColorsRecord splitMenuColors = new EscherSplitMenuColorsRecord();

        dggContainer.setRecordId(EscherContainerRecord.DGG_CONTAINER);
        dggContainer.setOptions(EscherContainerRecord.DGG_CONTAINER);

        EscherContainerRecord bStoreContainer = new EscherContainerRecord();
        bStoreContainer.setRecordId(EscherContainerRecord.BSTORE_CONTAINER);
        bStoreContainer.setOptions((short) (0xF));

        opt.setRecordId(EscherOptRecord.RECORD_ID);
        opt.addEscherProperty(new EscherBoolProperty(EscherPropertyTypes.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 524296));
        opt.addEscherProperty(new EscherRGBProperty(EscherPropertyTypes.FILL__FILLCOLOR, 0x08000041));
        opt.addEscherProperty(new EscherRGBProperty(EscherPropertyTypes.LINESTYLE__COLOR, 134217792));
        splitMenuColors.setRecordId(EscherSplitMenuColorsRecord.RECORD_ID);
        splitMenuColors.setOptions((short) 0x0040);
        splitMenuColors.setColor1(0x0800000D);
        splitMenuColors.setColor2(0x0800000C);
        splitMenuColors.setColor3(0x08000017);
        splitMenuColors.setColor4(0x100000F7);

        dggContainer.addChildRecord(dgg);
        dggContainer.addChildRecord(bStoreContainer);
        dggContainer.addChildRecord(opt);
        dggContainer.addChildRecord(splitMenuColors);

        DrawingGroupRecord drawingGroup = new DrawingGroupRecord();
        drawingGroup.addEscherRecord(dggContainer);
        return drawingGroup;
    }

    private static EscherDggRecord createEscherDggRecord() {
        EscherDggRecord dgg = new EscherDggRecord();
        dgg.setRecordId(EscherDggRecord.RECORD_ID);
        dgg.setShapeIdMax(1024);
        return dgg;
    }
    
    /**
     * Clears the cached list of drawing groups
     */
    public void clearDrawingGroups() {
        drawingGroups.clear(); 
    }

    /**
     * Creates a new drawing group 
     *
     * @return a new drawing group
     */
    public EscherDgRecord createDgRecord() {
        EscherDgRecord dg = new EscherDgRecord();
        dg.setRecordId( EscherDgRecord.RECORD_ID );
        short dgId = findNewDrawingGroupId();
        dg.setOptions( (short) ( dgId << 4 ) );
        dg.setNumShapes( 0 );
        dg.setLastMSOSPID( -1 );
        drawingGroups.add(dg);
        dgg.addCluster( dgId, 0 );
        dgg.setDrawingsSaved( dgg.getDrawingsSaved() + 1 );
        return dg;
    }

    /**
     * Allocates new shape id for the drawing group
     *
     * @param dg the EscherDgRecord which receives the new shape
     *
     * @return a new shape id.
     */
    public int allocateShapeId(EscherDgRecord dg) {
        return dgg.allocateShapeId(dg, true);
    }
    
    /**
     * Finds the next available (1 based) drawing group id
     * 
     * @return the next available drawing group id
     */
    public short findNewDrawingGroupId() {
        return dgg.findNewDrawingGroupId();
    }

    /**
     * Returns the drawing group container record
     *
     * @return the drawing group container record
     */
    public EscherDggRecord getDgg() {
        return dgg;
    }

    /**
     * Increment the drawing counter
     */
    public void incrementDrawingsSaved(){
        dgg.setDrawingsSaved(dgg.getDrawingsSaved()+1);
    }

    /**
     * Gets the record at the given 1-based index.
     * @param index 1-based index of the picture.
     * @return Record at the given index.
     * @throws IndexOutOfBoundsException if the index is larger than the number of available records.
     */
    EscherBSERecord getBSERecord(int index) {
        return (EscherBSERecord) bStoreContainer.getChild(index - 1);
    }

    /**
     * Allocates a new, empty picture in the workbook.
     * <p>
     * A picture is allocated by adding a {@link EscherBSERecord} for the picture to the workbook. The caller must
     * provide the picture data via {@link HSSFPictureData#setData(byte[])}.
     *
     * @param format One of the image format constants {@link Workbook#PICTURE_TYPE_EMF},
     *               {@link Workbook#PICTURE_TYPE_WMF}, {@link Workbook#PICTURE_TYPE_PICT},
     *               {@link Workbook#PICTURE_TYPE_JPEG}, {@link Workbook#PICTURE_TYPE_PNG}, or
     *               {@link Workbook#PICTURE_TYPE_DIB}.
     * @return Newly allocated picture without any image data.
     */
    @Internal
    public HSSFPictureData allocatePicture(int format) {
        EscherBlipRecord blipRecord;
        short escherTag;

        if (format == PICTURE_TYPE_EMF || format == PICTURE_TYPE_WMF) {
            blipRecord = new EscherMetafileBlip();
            escherTag = 0;
        } else {
            blipRecord = new EscherBitmapBlip();
            escherTag = (short) 0xFF;
        }

        blipRecord.setRecordId((short) (EscherBlipRecord.RECORD_ID_START + format));
        switch (format) {
            case PICTURE_TYPE_EMF:
                blipRecord.setOptions(HSSFPictureData.MSOBI_EMF);
                break;
            case PICTURE_TYPE_WMF:
                blipRecord.setOptions(HSSFPictureData.MSOBI_WMF);
                break;
            case PICTURE_TYPE_PICT:
                blipRecord.setOptions(HSSFPictureData.MSOBI_PICT);
                break;
            case PICTURE_TYPE_PNG:
                blipRecord.setOptions(HSSFPictureData.MSOBI_PNG);
                break;
            case PICTURE_TYPE_JPEG:
                blipRecord.setOptions(HSSFPictureData.MSOBI_JPEG);
                break;
            case PICTURE_TYPE_DIB:
                blipRecord.setOptions(HSSFPictureData.MSOBI_DIB);
                break;
            default:
                throw new IllegalStateException("Unexpected picture format: " + format);
        }

        EscherBSERecord r = new EscherBSERecord();
        r.setOptions((short) (0x0002 | (format << 4)));
        r.setBlipTypeMacOS((byte) format);
        r.setBlipTypeWin32((byte) format);
        r.setTag(escherTag);
        r.setBlipRecord(blipRecord);

        addBSERecord(r);
        return new HSSFPictureData(r);
    }

    /**
     * Returns the pictures in the workbook associated with this {@link DrawingManager2}.
     * <p>
     * The pictures are ordered as they are stored in the worksheet.
     *
     * @return Pictures in the workbook.
     */
    @Internal
    public List<HSSFPictureData> getAllPictures() {
        ArrayList<HSSFPictureData> pictures = new ArrayList<>(bStoreContainer.getChildCount());
        for (EscherRecord record : bStoreContainer) {
            EscherBSERecord bse = (EscherBSERecord) record;

            // We sometimes encounter sheets that have invalid BSE records
            EscherBlipRecord blip = bse.getBlipRecord();
            if (blip != null) {
                pictures.add(new HSSFPictureData(bse));
            } else {
                /** See {@link EscherBSERecord} */
                LOGGER.atDebug().log("Encountered BSE record without a BLIP. BSE records must have a BLIP according to the specification.");
            }
        }
        return pictures;
    }

    /**
     * Returns the number of pictures in the workbook associated with this {@link DrawingManager2}.
     *
     * @return Number of pictures in the workbook.
     */
    @Internal
    public int getPictureCount() {
        return bStoreContainer.getChildCount();
    }

    /**
     * Appends the given record to the {@link #bStoreContainer}.
     */
    private void addBSERecord(EscherBSERecord bse) {
        Objects.requireNonNull(bse);
        assert bse.getBlipRecord() != null;
        int index = bStoreContainer.getChildCount() + 1;
        bStoreContainer.setOptions((short) ((index << 4) | 0xF));
        bStoreContainer.addChildRecord(bse);
    }
}
