package pers.lzy.template.word.common;


import javax.annotation.Nonnull;
import java.io.Serializable;
import java.util.Iterator;
import java.util.Map;
import java.util.WeakHashMap;
import java.util.concurrent.locks.ReentrantReadWriteLock;

/**
 * 简单缓存，无超时实现，默认使用{@link WeakHashMap}实现缓存自动清理
 *	（来自 hutool工具类）
 * @param <K> 键类型
 * @param <V> 值类型
 */
public class SimpleCache<K, V> implements Iterable<Map.Entry<K, V>>, Serializable {
	private static final long serialVersionUID = 1L;

	/**
	 * 池
	 */
	private final Map<K, V> cache;
	// 乐观读写锁
	private final ReentrantReadWriteLock lock = new ReentrantReadWriteLock();

	/**
	 * 构造，默认使用{@link WeakHashMap}实现缓存自动清理
	 */
	public SimpleCache() {
		this(new WeakHashMap<>());
	}

	/**
	 * 构造
	 * <p>
	 * 通过自定义Map初始化，可以自定义缓存实现。<br>
	 * 比如使用{@link WeakHashMap}则会自动清理key，使用HashMap则不会清理<br>
	 * 同时，传入的Map对象也可以自带初始化的键值对，防止在get时创建
	 * </p>
	 *
	 * @param initMap 初始Map，用于定义Map类型
	 */
	public SimpleCache(Map<K, V> initMap) {
		this.cache = initMap;
	}

	/**
	 * 从缓存池中查找值
	 *
	 * @param key 键
	 * @return 值
	 */
	public V get(K key) {
		lock.readLock().lock();
		try {
			return cache.get(key);
		} finally {
			lock.readLock().unlock();
		}
	}

	/**
	 * 从缓存中获得对象，当对象不在缓存中或已经过期返回Func0回调产生的对象
	 *
	 * @param key      键
	 * @param supplier 如果不存在回调方法，用于生产值对象
	 * @return 值对象
	 */
	public V get(K key, Func0<V> supplier) {
		V v = get(key);

		if(null == v && null != supplier){
			lock.writeLock().lock();
			try{
				v = cache.get(key);
				// 双重检查，防止在竞争锁的过程中已经有其它线程写入
				if (null == v) {
					try {
						v = supplier.call();
					} catch (Exception e) {
						throw new RuntimeException(e);
					}
					cache.put(key, v);
				}
			} finally{
				lock.writeLock().unlock();
			}
		}

		return v;
	}

	/**
	 * 放入缓存
	 *
	 * @param key   键
	 * @param value 值
	 * @return 值
	 */
	public V put(K key, V value) {
		// 独占写锁
		lock.writeLock().lock();
		try {
			cache.put(key, value);
		} finally {
			lock.writeLock().unlock();
		}
		return value;
	}

	/**
	 * 移除缓存
	 *
	 * @param key 键
	 * @return 移除的值
	 */
	public V remove(K key) {
		// 独占写锁
		lock.writeLock().lock();
		try {
			return cache.remove(key);
		} finally {
			lock.writeLock().unlock();
		}
	}

	/**
	 * 清空缓存池
	 */
	public void clear() {
		// 独占写锁
		lock.writeLock().lock();
		try {
			this.cache.clear();
		} finally {
			lock.writeLock().unlock();
		}
	}

	@Override
	@Nonnull
	public Iterator<Map.Entry<K, V>> iterator() {
		return this.cache.entrySet().iterator();
	}


	/**
	 * 无参数的函数对象<br>
	 * 接口灵感来自于<a href="http://actframework.org/">ActFramework</a><br>
	 * 一个函数接口代表一个一个函数，用于包装一个函数为对象<br>
	 * 在JDK8之前，Java的函数并不能作为参数传递，也不能作为返回值存在，此接口用于将一个函数包装成为一个对象，从而传递对象
	 *
	 * @author Looly
	 *
	 * @param <R> 返回值类型
	 * @since 4.5.2
	 */
	@FunctionalInterface
	public interface Func0<R> {
		/**
		 * 执行函数
		 *
		 * @return 函数执行结果
		 * @throws Exception 自定义异常
		 */
		R call() throws Exception;

		/**
		 * 执行函数，异常包装为RuntimeException
		 *
		 * @return 函数执行结果
		 * @since 5.3.6
		 */
		default R callWithRuntimeException(){
			try {
				return call();
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}
	}
}