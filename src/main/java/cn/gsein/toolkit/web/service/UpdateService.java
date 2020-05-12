package cn.gsein.toolkit.web.service;

/**
 * @author G.Seinfeld
 * @date 2018/5/28
 */
public interface UpdateService<T> {
    /**
     * 变更实体信息
     *
     * @param t 实体的实例
     * @return 变更成功返回1，否则返回0
     */
    Integer updateEntity(T t);
}
