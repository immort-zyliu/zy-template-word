package pers.lzy.template.word.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import pers.lzy.template.word.anno.HandlerOrder;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.ServiceLoader;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/7/11  15:47
 */
public class SpiLoader {


    private final static Logger logger = LoggerFactory.getLogger(SpiLoader.class);

    private static final Map<String, ServiceLoader<?>> SERVICE_LOADER_MAP = new ConcurrentHashMap<>();

    /**
     * Load the sorted SPI instance list for provided SPI interface.
     * <p>
     * Note: each call return same instances.
     *
     * @param clazz class of the SPI
     * @param <T>   SPI type
     * @return sorted SPI instance list
     */
    public static <T> List<T> loadInstanceListSorted(Class<T> clazz) {
        try {
            String key = clazz.getName();
            // Not thread-safe, as it's expected to be resolved in a thread-safe context.
            @SuppressWarnings("unchecked")
            ServiceLoader<T> serviceLoader = (ServiceLoader<T>) SERVICE_LOADER_MAP.get(key);
            if (serviceLoader == null) {
                serviceLoader = ServiceLoaderUtil.getServiceLoader(clazz);
                SERVICE_LOADER_MAP.put(key, serviceLoader);
            }

            List<SpiOrderWrapper<T>> orderWrappers = new ArrayList<>();
            for (T spi : serviceLoader) {
                int order = SpiOrderResolver.resolveOrder(spi);
                // Since SPI is lazy initialized in ServiceLoader, we use online sort algorithm here.
                SpiOrderResolver.insertSorted(orderWrappers, spi, order);
                logger.info("[SpiLoader] Found {} SPI: {} with order {}", clazz.getSimpleName(),
                        spi.getClass().getCanonicalName(), order);
            }
            List<T> list = new ArrayList<>(orderWrappers.size());
            for (SpiOrderWrapper<T> orderWrapper : orderWrappers) {
                list.add(orderWrapper.spi);
            }
            return list;
        } catch (Throwable t) {
            logger.error("[SpiLoader] ERROR: loadInstanceListSorted failed", t);
        }
        return new ArrayList<>();
    }


    private static class SpiOrderResolver {
        private static <T> void insertSorted(List<SpiOrderWrapper<T>> list, T spi, int order) {
            int idx = 0;
            for (; idx < list.size(); idx++) {
                if (list.get(idx).getOrder() > order) {
                    break;
                }
            }
            list.add(idx, new SpiOrderWrapper<>(order, spi));
        }

        private static <T> int resolveOrder(T spi) {
            if (!spi.getClass().isAnnotationPresent(HandlerOrder.class)) {
                // Lowest precedence by default.
                return HandlerOrder.LOWEST_PRECEDENCE;
            } else {
                return spi.getClass().getAnnotation(HandlerOrder.class).value();
            }
        }
    }

    private static class SpiOrderWrapper<T> {
        private final int order;
        private final T spi;

        SpiOrderWrapper(int order, T spi) {
            this.order = order;
            this.spi = spi;
        }

        int getOrder() {
            return order;
        }

        T getSpi() {
            return spi;
        }
    }

}
