package org.apache.poi.openxml4j.opc;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

public class TestZipPackage {

    @Test
    void testDefaultDirectory(@TempDir Path tempDir) {
        try {
            assertNull(ZipPackage.getTempDirectory());
            ZipPackage.setTempDirectory(tempDir);
            assertEquals(tempDir, ZipPackage.getTempDirectory());
        } finally {
            ZipPackage.setTempDirectory(null);
            assertNull(ZipPackage.getTempDirectory());
        }
    }
}
