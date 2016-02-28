package org.life.Core.TypeFlag;

import org.life.Core.TypeFlag.AbstractFlag.Flag;

import java.io.FileInputStream;
import java.io.IOException;

public class HSSFFlag implements Flag{
    private FileInputStream fileStream;

    public HSSFFlag() {}
    public HSSFFlag(FileInputStream fileStream) {
        this.fileStream = fileStream;
    }

    @Override
    public FileInputStream getStream()
    {
        return fileStream;
    }

    @Override
    public void close()
    {
        try
        {
            fileStream.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
            throw new RuntimeException(String
                    .format("%s: close stream failure (Method: close).", getClass().getName()));
        }
    }
}
