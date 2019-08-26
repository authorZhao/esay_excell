package com.git.poi.validator;

import com.git.poi.util.RegexUtils;
import java.util.function.Predicate;

/**
 * 校验的方法
 */
public class Check {

    public static final Predicate<String> PREDICAT_EEMAIL = RegexUtils::isEmail;
    public static final Predicate<String> PREDICAT_EMOBILE = RegexUtils::isMobileNum;

    public static Boolean isEmail(String email) {
        return PREDICAT_EEMAIL.test(email);
    }
    public static Boolean isMobileNum(String mobileNum) {
        return PREDICAT_EMOBILE.test(mobileNum);
    }
}
