package cn.gsein.toolkit.web.service;

/**
 * @author G.Seinfeld
 * @date 2018/5/28
 */
public interface DeleteService<T> {
    /**
     * 删除提供的实体信息，要求实现该方法时执行软删除，即在数据库层执行update操作，将is_deleted字段的值设置为1
     *
     * @param t 实体的实例
     * @return 删除成功返回1，否则返回0
     */
    Integer deleteEntity(T t);

    /**
     * 根据主键删除实体信息，要求实现该方法时执行软删除，即在数据库层执行update操作，将is_deleted字段的值设置为1
     *
     * @param key 实体的主键，一般为String或者int类型
     * @return 删除成功返回1，否则返回0
     */
    Integer deleteEntityByKey(Object key);
}
