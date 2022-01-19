package cn.gsein.toolkit.utils;

import org.apache.commons.lang3.RandomStringUtils;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.UUID;
import java.util.stream.Collectors;

public final class GseinStringUtils {

    private GseinStringUtils() {
    }

    /**
     * 合并指定符号分隔的字符串
     *
     * @param strs      要合并的字符串
     * @param trim      是否去掉两端空格
     * @param delimiter 分隔符
     * @return 合并后的字符串
     */
    public static String mergeSymbolSeparatedString(String delimiter, boolean trim, String... strs) {
        return Arrays.stream(strs)
                .filter(str -> str != null && str.length() > 0)
                .flatMap(str -> Arrays.stream(str.split(delimiter)))
                .map(str -> trim ? str.trim() : str)
                .distinct()
                .collect(Collectors.joining(delimiter));
    }

    /**
     * 生成随机UUID字符串，去掉-号
     *
     * @return UUID字符串
     */
    public static String generateUUID() {
        return UUID.randomUUID().toString().replace("-", "");
    }

    /**
     * 生成流水号 格式 指定前缀-yyyyMMddHHmmss格式的时间-四位随机数字字母
     *
     * @param prefix 指定前缀
     * @return 流水号
     */
    public static String generateBillCode(String prefix) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String time = formatter.format(LocalDateTime.now());

        String random = RandomStringUtils.randomAlphanumeric(4);

        return prefix + "-" + time + "-" + random;
    }

    public static void main(String[] args) {
        String s = mergeSymbolSeparatedString(",", true, "1, 2, 3", "4, 3, 5", "", null, "5, 6, 7");
        System.out.println(s);
    }
}
