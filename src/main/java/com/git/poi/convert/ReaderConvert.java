package com.git.poi.convert;


import com.git.poi.exception.ConvertException;

public interface ReaderConvert {
    Object convert(Object obj) throws ConvertException;
}
