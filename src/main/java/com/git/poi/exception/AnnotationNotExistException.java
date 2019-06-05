package com.git.poi.exception;

public class AnnotationNotExistException extends RuntimeException {

    private static final long serialVersionUID = -1169100027771948958L;

    public AnnotationNotExistException() {
        super();
    }

    public AnnotationNotExistException(String message) {
        super(message);
    }

    public AnnotationNotExistException(String message, Throwable cause) {
        super(message, cause);
    }

    public AnnotationNotExistException(Throwable cause) {
        super(cause);
    }

    protected AnnotationNotExistException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
