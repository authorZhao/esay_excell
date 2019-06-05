package com.git.poi.exception;

public class ClassMappingException extends RuntimeException {

    private static final long serialVersionUID = -1169100027771948958L;

    public ClassMappingException() {
        super();
    }

    public ClassMappingException(String message) {
        super(message);
    }

    public ClassMappingException(String message, Throwable cause) {
        super(message, cause);
    }

    public ClassMappingException(Throwable cause) {
        super(cause);
    }

    protected ClassMappingException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
