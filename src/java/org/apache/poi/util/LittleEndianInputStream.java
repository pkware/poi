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

package org.apache.poi.util;

import java.io.FilterInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Wraps an {@link InputStream} providing {@link LittleEndianInput}<p/>
 *
 * This class does not buffer any input, so the stream read position maintained
 * by this class is consistent with that of the inner stream.
 *
 * @author Josh Micich
 */
public class LittleEndianInputStream extends FilterInputStream implements LittleEndianInput {
	public LittleEndianInputStream(InputStream is) {
		super(is);
	}
	
	@Override
    public int available() {
		try {
			return super.available();
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}
	
	@Override
    public byte readByte() {
		return (byte)readUByte();
	}
	
	@Override
    public int readUByte() {
		byte buf[] = new byte[1];
		try {
			checkEOF(read(buf), 1);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
		return LittleEndian.getUByte(buf);
	}
	
	@Override
    public double readDouble() {
		return Double.longBitsToDouble(readLong());
	}
	
	@Override
    public int readInt() {
	    byte buf[] = new byte[LittleEndianConsts.INT_SIZE];
		try {
		    checkEOF(read(buf), buf.length);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
		return LittleEndian.getInt(buf);
	}
	
    /**
     * get an unsigned int value from an InputStream
     * 
     * @return the unsigned int (32-bit) value
     * @exception RuntimeException
     *                wraps any IOException thrown from reading the stream.
     */
    public long readUInt() {
       long retNum = readInt();
       return retNum & 0x00FFFFFFFFL;
    }
	
	@Override
    public long readLong() {
		byte buf[] = new byte[LittleEndianConsts.LONG_SIZE];
		try {
		    checkEOF(read(buf), LittleEndianConsts.LONG_SIZE);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
		return LittleEndian.getLong(buf);
	}
	
	@Override
    public short readShort() {
		return (short)readUShort();
	}
	
	@Override
    public int readUShort() {
		byte buf[] = new byte[LittleEndianConsts.SHORT_SIZE];
		try {
		    checkEOF(read(buf), LittleEndianConsts.SHORT_SIZE);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
		return LittleEndian.getUShort(buf);
	}
	
	private static void checkEOF(int actualBytes, int expectedBytes) {
		if (expectedBytes != 0 && (actualBytes == -1 || actualBytes != expectedBytes)) {
			throw new RuntimeException("Unexpected end-of-file");
		}
	}

	@Override
    public void readFully(byte[] buf) {
		readFully(buf, 0, buf.length);
	}

	@Override
    public void readFully(byte[] buf, int off, int len) {
	    try {
	        checkEOF(read(buf, off, len), len);
	    } catch (IOException e) {
            throw new RuntimeException(e);
        }
	}

    @Override
    public void readPlain(byte[] buf, int off, int len) {
        readFully(buf, off, len);
    }
}
