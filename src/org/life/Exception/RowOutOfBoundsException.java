package org.life.Exception;

public class RowOutOfBoundsException extends RuntimeException {
    public RowOutOfBoundsException() {
    }

    public RowOutOfBoundsException(String message) {
        super(message);
    }

    public RowOutOfBoundsException(String message, Throwable cause) {
        super(message, cause);
    }

    public RowOutOfBoundsException(Throwable cause) {
        super(cause);
    }

    public RowOutOfBoundsException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
