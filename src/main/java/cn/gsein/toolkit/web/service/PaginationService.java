package cn.gsein.toolkit.web.service;

import cn.gsein.toolkit.web.model.PageRequestData;
import com.github.pagehelper.PageInfo;

import java.util.List;
import java.util.Map;

/**
 * @author G.Seinfeld
 * @date 2018/5/28
 */
public interface PaginationService<T> {
    /**
     * 实体分页信息
     */
    List<T> pageEntity(Integer pageNum, Integer pageSize);

    /**
     * 实体分页信息
     */
    PageInfo<T> pageEntitys(Integer pageNum, Integer pageSize, String orderBy);

    /**
     * 实体分页信息
     *
     * @param pageNum          页数
     * @param pageSize         每页行数
     * @param searchConditions 搜索条件
     * @param orderBy          排序sql拼接
     * @return 分布信息
     */
    PageInfo<T> pageInfo(Integer pageNum, Integer pageSize, Map<String, Object> searchConditions, String orderBy);

    /**
     * 获取某实体分页信息
     *
     * @param pageRequestData           封装的分页请求数据
     * @param needsAuthorizationControl 是否需要进行权限控制（按部门）
     * @return 实体分页信息
     */
    PageInfo<T> pageInfo(PageRequestData pageRequestData, boolean needsAuthorizationControl);
}
