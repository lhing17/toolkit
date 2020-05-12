package cn.gsein.toolkit.web.service;

/**
 * @author G.Seinfeld
 * @date 2018/5/28
 */
public interface SaveService<T> {
    /**
     * 保存实体信息
     *
     * @param t 实体的实例
     * @return 保存成功返回1，否则返回0
     */
    Integer saveEntity(T t);
}
