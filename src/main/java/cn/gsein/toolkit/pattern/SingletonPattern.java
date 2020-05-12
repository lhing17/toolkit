package cn.gsein.toolkit.pattern;

/**
 * @author G. Seinfeld
 */
public class SingletonPattern {

}

/**
 * 饿汉式单例类。在类初始化时，已经自行实例化
 */
class HungerSingleton {
    private HungerSingleton() {
    }

    @SuppressWarnings("InstantiationOfUtilityClass")
    private static final HungerSingleton SINGLE = new HungerSingleton();

    // 静态工厂方法
    public static HungerSingleton getInstance() {
        return SINGLE;
    }
}

/**
 * 懒汉式单例类.在第一次调用的时候实例化自己
 */
class LazySingleton {
    private LazySingleton() {
    }

    private static LazySingleton singleton = null;

    /**
     * 静态工厂方法，加锁保证线程安全
     */
    public static LazySingleton getSingleton() {
        if (singleton == null) {
            synchronized (HungerSingleton.class) {
                if (singleton == null) {
                    //noinspection InstantiationOfUtilityClass
                    singleton = new LazySingleton();
                }
            }
        }
        return singleton;
    }
}

/**
 * 枚举单例
 * 单元素的枚举类型已经成为实现Singleton的最佳方法。
 * 优点：自由序列化，线程安全，保证单例
 */
class Resource {
}

enum SomeThing {
    /**
     * 单例模式的实例
     */
    INSTANCE;
    private final Resource instance;

    SomeThing() {
        instance = new Resource();
    }

    public Resource getInstance() {
        return instance;
    }
}
