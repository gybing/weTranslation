package org.life.Exception;

public class FlagException extends RuntimeException {
    public FlagException() {
    }

    public FlagException(String message) {
        super(message);
    }

    public FlagException(String message, Throwable cause) {
        super(message, cause);
    }

    public FlagException(Throwable cause) {
        super(cause);
    }

    public FlagException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
