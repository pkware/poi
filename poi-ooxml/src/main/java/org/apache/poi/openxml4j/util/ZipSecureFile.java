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

package org.apache.poi.openxml4j.util;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.EnumSet;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * This class wraps a {@link ZipFile} in order to check the
 * entries for <a href="https://en.wikipedia.org/wiki/Zip_bomb">zip bombs</a>
 * while reading the archive.<p>
 *
 * The alert limits can be globally defined via {@link #setMaxEntrySize(long)}
 * and {@link #setMinInflateRatio(double)}.
 */
public class ZipSecureFile extends ZipFile {
    private static final Logger LOG = LogManager.getLogger(ZipSecureFile.class);
    /* package */ static double MIN_INFLATE_RATIO = 0.01d;
    /* package */ static final long DEFAULT_MAX_ENTRY_SIZE = 0xFFFFFFFFL;
    /* package */ static long MAX_ENTRY_SIZE = DEFAULT_MAX_ENTRY_SIZE;

    // The maximum chars of extracted text
    /* package */ static final long DEFAULT_MAX_TEXT_SIZE = 10*1024*1024L;
    private static long MAX_TEXT_SIZE = DEFAULT_MAX_TEXT_SIZE;

    private final String fileName;

    /**
     * Sets the ratio between de- and inflated bytes to detect zipbomb.
     * It defaults to 1% (= 0.01d), i.e. when the compression is better than
     * 1% for any given read package part, the parsing will fail indicating a 
     * Zip-Bomb.
     *
     * @param ratio the ratio between de- and inflated bytes to detect zipbomb
     */
    public static void setMinInflateRatio(double ratio) {
        MIN_INFLATE_RATIO = ratio;
    }
    
    /**
     * Returns the current minimum compression rate that is used.
     * 
     * See setMinInflateRatio() for details.
     *
     * @return The min accepted compression-ratio.  
     */
    public static double getMinInflateRatio() {
        return MIN_INFLATE_RATIO;
    }

    /**
     * Sets the maximum file size of a single zip entry. It defaults to 4GB,
     * i.e. the 32-bit zip format maximum.
     * 
     * This can be used to limit memory consumption and protect against 
     * security vulnerabilities when documents are provided by users.
     *
     * @param maxEntrySize the max. file size of a single zip entry
     * @throws IllegalArgumentException for negative maxEntrySize
     */
    public static void setMaxEntrySize(long maxEntrySize) {
        if (maxEntrySize < 0) {
            throw new IllegalArgumentException("Max entry size must be greater than or equal to zero");
        } else if (maxEntrySize > DEFAULT_MAX_ENTRY_SIZE) {
            LOG.atWarn().log("setting max entry size greater than 4Gb can be risky; set to " + maxEntrySize + " bytes");
        }
        MAX_ENTRY_SIZE = maxEntrySize;
    }

    /**
     * Returns the current maximum allowed uncompressed file size.
     * 
     * See setMaxEntrySize() for details.
     *
     * @return The max accepted uncompressed file size. 
     */
    public static long getMaxEntrySize() {
        return MAX_ENTRY_SIZE;
    }

    /**
     * Sets the maximum number of characters of text that are
     * extracted before an exception is thrown during extracting
     * text from documents.
     * 
     * This can be used to limit memory consumption and protect against 
     * security vulnerabilities when documents are provided by users.
     *
     * @param maxTextSize the max. file size of a single zip entry
     * @throws IllegalArgumentException for negative maxTextSize
     */
    public static void setMaxTextSize(long maxTextSize) {
        if (maxTextSize < 0) {
            throw new IllegalArgumentException("Max text size must be greater than or equal to zero");
        }else if (maxTextSize > DEFAULT_MAX_TEXT_SIZE) {
            LOG.atWarn().log("setting max text size greater than " + DEFAULT_MAX_TEXT_SIZE + " can be risky; set to " + maxTextSize + " chars");
        }
        MAX_TEXT_SIZE = maxTextSize;
    }

    /**
     * Returns the current maximum allowed text size.
     * 
     * @return The max accepted text size.
     * @see #setMaxTextSize(long)
     */
    public static long getMaxTextSize() {
        return MAX_TEXT_SIZE;
    }

    public ZipSecureFile(File file) throws IOException {
        this(file.toPath());
    }

    public ZipSecureFile(Path path) throws IOException {
        super(
            Files.newByteChannel(path, EnumSet.of(StandardOpenOption.READ)),
            path.toAbsolutePath().toString(),
            StandardCharsets.UTF_8.name(),
            true,
            false
        );
        this.fileName = path.toAbsolutePath().toString();
    }

    public ZipSecureFile(String name) throws IOException {
        this(Paths.get(name));
    }

    /**
     * Returns an input stream for reading the contents of the specified
     * zip file entry.
     *
     * <p> Closing this ZIP file will, in turn, close all input
     * streams that have been returned by invocations of this method.
     *
     * @param entry the zip file entry
     * @return the input stream for reading the contents of the specified
     * zip file entry.
     * @throws IOException if an I/O error has occurred
     * @throws IllegalStateException if the zip file has been closed
     */
    @Override
    public ZipArchiveThresholdInputStream getInputStream(ZipArchiveEntry entry) throws IOException {
        ZipArchiveThresholdInputStream zatis = new ZipArchiveThresholdInputStream(super.getInputStream(entry));
        zatis.setEntry(entry);
        return zatis;
    }

    /**
     * Returns the path name of the ZIP file.
     * @return the path name of the ZIP file
     */
    public String getName() {
        return fileName;
    }
}
