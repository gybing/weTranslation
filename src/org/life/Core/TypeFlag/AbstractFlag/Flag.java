package org.life.Core.TypeFlag.AbstractFlag;

import java.io.FileInputStream;

public interface Flag{

    FileInputStream getStream();

    void close();
}
