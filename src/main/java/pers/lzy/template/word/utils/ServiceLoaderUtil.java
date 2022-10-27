package pers.lzy.template.word.utils;

import java.util.ServiceLoader;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/7/11  16:01
 */
public class ServiceLoaderUtil {

    static {
        initConfig();
    }

    private static String SPI_CLASSLOADER = "default";

    private static final String CLASSLOADER_DEFAULT = "default";
    private static final String CLASSLOADER_CONTEXT = "context";

    public static <S> ServiceLoader<S> getServiceLoader(Class<S> clazz) {
        if (shouldUseContextClassloader()) {
            return ServiceLoader.load(clazz);
        } else {
            return ServiceLoader.load(clazz, clazz.getClassLoader());
        }
    }

    public static boolean shouldUseContextClassloader() {
        String classloaderConf = SPI_CLASSLOADER;
        return CLASSLOADER_CONTEXT.equalsIgnoreCase(classloaderConf);
    }

    private static void initConfig() {
        // 有机会在扩展吧....
        SPI_CLASSLOADER = "default";
    }
}
