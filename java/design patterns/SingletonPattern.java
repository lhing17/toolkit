
public class SingletonPattern {

}

/**
 * 饿汉式单例类。在类初始化时，已经自行实例化   
 */
class Singleton {
    private Singleton() {
    }

    private static final Singleton single = new Singleton();

    // 静态工厂方法   
    public static Singleton getInstance() {
        return single;
    }
}

/**
 * 懒汉式单例类.在第一次调用的时候实例化自己
 */
class Singleton {
    private Singleton() {
    }

    private static Singleton single = null;

    // 静态工厂方法，加锁保证线程安全
    public static Singleton getSingleton() {    
        if (singleton == null) {    
            synchronized (Singleton.class) {    
                if (singleton == null) {    
                    singleton = new Singleton();    
                }    
            }    
        }    
    }
}

/**
 * 枚举单例 
 * 单元素的枚举类型已经成为实现Singleton的最佳方法。
 * 优点：自由序列化，线程安全，保证单例
 */
class Resource{
}

public enum SomeThing {
    INSTANCE;
    private Resource instance;
    SomeThing() {
        instance = new Resource();
    }
    public Resource getInstance() {
        return instance;
    }
}