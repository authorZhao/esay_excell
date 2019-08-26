package com.git.poi.util;

import org.apache.commons.lang3.StringUtils;


public class RegexUtils {

    /**
     * 用户名是否符合规范（^[\u4E00-\u9FA5A-Za-z0-9_]+$）
     * @return
     */
    public static boolean isValidUsername(String username) {
        if (StringUtils.isBlank(username)) {
            return false;
        }

        return username.matches("^[\u4E00-\u9FA5A-Za-z0-9_]{3,10}$");
    }

    /**
     * URL是否规范（^((http|ftp|https)://)(([a-zA-Z0-9\\._-]+\\.[a-zA-Z]{2,6})|([0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}))(:[0-9]{1,4})*(/[a-zA-Z0-9\\&amp;%_\\./-~-]*)?$）
     * @param url
     * @return
     */
    public static boolean isValidUrl(String url) {
        if (StringUtils.isBlank(url)) {
            return false;
        }
        return url.matches("^((http|ftp|https)://)(([a-zA-Z0-9\\._-]+\\.[a-zA-Z]{2,6})|([0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}))(:[0-9]{1,4})*(/[a-zA-Z0-9\\&amp;%_\\./-~-]*)?$");
    }

    /**
     * 密码是否符合规范（[a-zA-Z\d]{6,20}）
     * @return
     */
    public static boolean isValidPassword(String password) {
        if (null == password) {
            return false;
        }

        return password.matches("^([^\\s'‘’]{6,20})$");
    }

    /**
     * 是否有效手机号码
     * @param mobileNum
     * @return
     */
    public static boolean isMobileNum(String mobileNum) {
        if (null == mobileNum) {
            return false;
        }

        return mobileNum.matches("^((13[0-9])|(14[4,7])|(15[^4,\\D])|(17[0-9])|(18[0-9]))(\\d{8})$");
    }

    /**
     * 是否有效邮箱
     * @param email
     * @return
     */
    public static boolean isEmail(String email) {
        if (null == email) {
            return false;
        }

        String regex = "^\\s*\\w+(?:\\.{0,1}[\\w-]+)*@[a-zA-Z0-9]+(?:[-.][a-zA-Z0-9]+)*\\.[a-zA-Z]+\\s*$";
        return email.matches(regex);
    }

    /**
     * 是否是QQ邮箱
     */
    public static boolean isQQEmail(String email){
        if(null == email)
            return false;

        return email.matches("^[\\s\\S]*@qq.com$");
    }

    /**
     * 是否数字(小数||整数)
     * @param number
     * @return
     */
    public static boolean isNumber(String number) {
        if (null == number) {
            return false;
        }

        return number.matches("^[+-]?(([1-9]{1}\\d*)|([0]{1}))(\\.(\\d)+)?$");
    }
    /**
     * 是否为不为手机号码的数字(小数||整数)
     * @param number
     * @return
     */
    public static boolean isNumberAndNotMobile(String number) {
        if (null == number) {
            return false;
        }
        boolean matches = number.matches("^((13[0-9])|(14[4,7])|(15[^4,\\D])|(17[0-9])|(18[0-9]))(\\d{8})$");
        if (matches){ //true则表示是手机号码
            return false;
        }
        return number.matches("^[+-]?(([1-9]{1}\\d*)|([0]{1}))(\\.(\\d)+)?$");
    }

    /**
     * 是否整数
     * @param number
     * @return
     */
    public static boolean isInt(String number) {
        if (null == number) {
            return false;
        }

        return number.matches("^[+-]?(([1-9]{1}\\d*)|([0]{1}))$");
    }

    /**
     * 是否正整数
     * @param number
     * @return
     */
    public static boolean isPositiveInt(String number) {
        if (null == number) {
            return false;
        }

        return number.matches("^[+-]?(([1-9]{1}\\d*)|([0]{1}))$");
    }

    /**
     * 是否日期yyyy-mm-dd(yyyy/mm/dd)
     * @param date
     * @return
     */
    public static boolean isDate(String date) {
        if (null == date) {
            return false;
        }
        return date.matches("^([1-2]\\d{3})[\\/|\\-](0?[1-9]|10|11|12)[\\/|\\-]([1-2]?[0-9]|0[1-9]|30|31)$");
    }

    /**
     * 逗号分隔的正则表达式
     * @param str
     * @return
     */
    public static String getCommaSparatedRegex(String str) {
        if (str == null) {
            return null;
        }

        return "^("+str+")|([\\s\\S]*,"+str+")|("+str+",[\\s\\S]*)|([\\s\\S]*,"+str+",[\\s\\S]*)$";
    }

    /**
     * 字符串包含
     * @return
     */
    public static boolean contains(String str, String regex) {
        if (str == null || regex == null) {
            return false;
        }

        return str.matches(regex);
    }

    /**
     * 是否为16-22位银行账号
     * @param bankAccount
     * @return
     */
    public static boolean isBankAccount(String bankAccount){
        if (null == bankAccount) {
            return false;
        }

        return bankAccount.matches("^\\d{16,22}$");
    }

    public static boolean isValidDomain(String domain) {
        if (StringUtils.isBlank(domain)) {
            return false;
        }
        return domain.matches("^[a-zA-Z0-9\u4e00-\u9fa5][-a-zA-Z0-9\u4e00-\u9fa5]{0,62}(\\.[a-zA-Z0-9\u4e00-\u9fa5][-a-zA-Z0-9\u4e00-\u9fa5]{0,62})+\\.?$");
    }

    /**
     * 是否是qq号码
     *
     * @param qq
     * @return
     * @description
     *
     * @author lhf
     * @createDate 2017年9月12日
     */
    public static boolean isQQnumber(String qq) {
        if (StringUtils.isBlank(qq)) {
            return false;
        }
        return qq.matches("^[1-9][0-9]{4,11}?$");
    }

    /**
     * 是否中文名字
     */
    public static boolean isUserName(String name){
        if (StringUtils.isBlank(name)){
            return false;
        }
        return name.matches("[\u4e00-\u9fa5]+");
    }

    public static void main(String[] args) {
        //验收邮箱是否合法
        boolean flag1 = isEmail("8@lomoo.com");
        System.out.println(flag1);

        //验证用户名是否合法
        boolean flag2 = isValidUsername("***&amp;霜花似雪&amp;***");
        System.out.println(flag2);

        boolean flag3 = isValidUrl("http://www.baidu");
        System.out.println(flag3);

        boolean flag4 = isValidPassword("6776tsgs188*&amp;^^%");
        System.out.println(flag4);

        String strCode = "18396676111";
        System.out.println(isMobileNum(strCode));
        if(isMobileNum(strCode)){
            System.out.println(strCode + "是合法的手机号");
        }else if(isQQnumber(strCode)){
            System.out.println(strCode + "是合法的QQ号");
        }else if(isPositiveInt(strCode)){
            System.out.println(strCode + "是正整数");
        }else{
            System.out.println("字符串");
        }

    }
}
