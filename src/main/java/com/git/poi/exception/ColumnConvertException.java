package com.git.poi.exception;

public class ColumnConvertException extends ConvertException {
    public ColumnConvertException() {
        super();
    }

    public ColumnConvertException(String message) {
        super(message);
    }

    public ColumnConvertException(String message, Throwable cause) {
        super(message, cause);
    }

    public ColumnConvertException(Throwable cause) {
        super(cause);
    }

    protected ColumnConvertException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
