package cn.iocoder.yudao.module.cg.service.cgtest;


import cn.iocoder.yudao.framework.common.pojo.PageResult;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgPagePeqVO;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgSaveReqVO;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgUpdateReqVO;
import cn.iocoder.yudao.module.cg.dal.dataobject.cgtest.CgTestDO;
import jakarta.validation.Valid;

import java.util.Collection;
import java.util.List;

/**
 * 后台 cg Service 接口
 */
public interface CgTestService {
    /**
     * 创建用户
     *
     * @param createReqVO 用户信息
     * @return 用户编号
     */
    Long createUser(@Valid CgSaveReqVO createReqVO);

    /**
     * 修改用户
     *
     * @param updateReqVO 用户信息
     */
    void updateUser(@Valid CgUpdateReqVO updateReqVO);

    /**
     * 删除用户
     *
     * @param id 用户编号
     */
    void deleteUser(Long id);

    void deleteUsers(List<Long> ids);

    /**
     * 通过用户名查询用户
     *
     * @param username 用户名
     * @return 用户对象信息
     */
    CgTestDO getUserByUsername(String username);

    /**
     * 通过用户 ID 查询用户
     *
     * @param id 用户ID
     * @return 用户对象信息
     */
    CgTestDO getUser(Long id);

    /**
     * 获得用户分页列表
     *
     * @param reqVO 分页条件
     * @return 分页列表
     */
    PageResult<CgTestDO> getUserPage(CgPagePeqVO reqVO);

    /**
     * 获得用户列表
     *
     * @param ids 用户编号数组
     * @return 用户列表
     */
    List<CgTestDO> getUserList(Collection<Long> ids);

    /**
     * 获得指定状态的用户们
     *
     * @param ageNumber 状态
     * @return 用户们
     */
    List<CgTestDO> getUserListByStatus(Integer ageNumber);
}
