package com.git.poi.validator;

import com.git.poi.util.RegexUtils;
import java.util.function.Predicate;

/**
 * 校验的方法
 */
public class Check {

    public static Boolean isPassCheck(String str,Predicate<String> predicate) {
        return predicate.test(str);
    }
}
