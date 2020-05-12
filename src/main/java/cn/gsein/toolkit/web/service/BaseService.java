package cn.gsein.toolkit.web.service;

/**
 * 基础Service类，各业务的Service类必须继承此接口
 *
 * @author G. Seinfeld
 * @date 2017/12/20
 */
public interface BaseService<T>
                extends QueryService<T>, PaginationService<T>, SaveService<T>, UpdateService<T>, DeleteService<T> {

}
