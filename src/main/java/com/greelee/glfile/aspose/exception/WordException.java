package com.greelee.glfile.aspose.exception;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/27
 * @describe:
 */
public class WordException extends RuntimeException {
    private static final long serialVersionUID = 4714583207238573668L;

    /**
     * Constructs a new runtime exception with the specified detail message.
     * The cause is not initialized, and may subsequently be initialized by a
     * call to {@link #initCause}.
     *
     * @param message the detail message. The detail message is saved for
     *                later retrieval by the {@link #getMessage()} method.
     */
    public WordException(String message) {
        super(message);
    }
}
