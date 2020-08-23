package com.dlw.architecture.office.exception;

/**
 * @author dengliwen
 * @date 2020/6/11
 * @desc office 操作异常
 * @since 4.0.0
 */
public class OfficeException extends Exception {

    public OfficeException(String message) {
        super(message);
    }

    public OfficeException(String message, Throwable cause) {
        super(message, cause);
    }
}
