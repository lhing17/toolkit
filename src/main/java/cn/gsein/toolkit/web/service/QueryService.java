package cn.gsein.toolkit.web.service;

import java.util.List;

/**
 * @author G.Seinfeld
 * @date 2018/5/28
 */
public interface QueryService<T> {
    /**
     * 根据主键获得实体信息
     *
     * @param key 实体的主键，一般为String或者int类型
     * @return 查询到的实例
     */
    T queryEntityByKey(Object key);

    /**
     * 查询所有实体信息
     */
    List<T> queryEntity(T t, String column);

    /**
     * 根据实体信息，查询实体列表信息
     */
    List<T> queryEntity(T t);
}
