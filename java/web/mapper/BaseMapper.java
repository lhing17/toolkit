package com.jweb.db.mapper;

import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;

/**
 * mybatis的基础Mapper类，基本的CRUD
 * 
 * @author G.Seinfeld
 * @date 2017/11/9
 */
@Component
public interface BaseMapper<T> {
    /**
     * 插入一条数据
     * 
     * @param t 要插入的数据
     * @return 插入成功返回1，否则返回0
     */
    int insert(T t);

    /**
     * 删除一条数据
     * 
     * @param t 要删除的数据
     * @return 删除成功返回1，否则返回0
     */
    int delete(T t);

    /**
     * 更新一条数据
     * 
     * @param t 要更新的数据
     * @return 更新成功返回1，否则返回0
     */
    int update(T t);

    /**
     * 根据主键删除一条数据
     * 
     * @param id 要删除数据的主键
     * @return 删除成功返回1，否则返回0
     */
    int deleteById(Integer id);

     /**
     * 根据主键查询到唯一一条数据
     * 
     * @param id 要查询数据的主键
     * @return 查询到的唯一一条实体数据
     */
    T queryById(Integer id);

    /**
     * 根据条件查询出某实体类数据的列表
     * 
     * @param conditions 查询条件
     * @return 实体数据列表
     */
    List<T> queryList(Map<String, Object> conditions);


}
