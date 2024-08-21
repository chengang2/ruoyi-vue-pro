package cn.iocoder.yudao.module.cg.dal.mysql.cgtest;


import cn.iocoder.yudao.framework.common.pojo.PageResult;
import cn.iocoder.yudao.framework.mybatis.core.mapper.BaseMapperX;
import cn.iocoder.yudao.framework.mybatis.core.query.LambdaQueryWrapperX;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgPagePeqVO;
import cn.iocoder.yudao.module.cg.dal.dataobject.cgtest.CgTestDO;
import org.apache.ibatis.annotations.Mapper;

import java.util.Collection;
import java.util.List;

@Mapper
public interface CgTestMapper extends BaseMapperX<CgTestDO> {
    default CgTestDO selectByUsername(String name) {
        return selectOne(CgTestDO::getName, name);
    }

    default CgTestDO selectByAgeNumber(Integer ageNumber) {
        return selectOne(CgTestDO::getAgeNumber, ageNumber);
    }

    default PageResult<CgTestDO> selectPage(CgPagePeqVO reqVO) {
        return selectPage(reqVO, new LambdaQueryWrapperX<CgTestDO>()
                .likeIfPresent(CgTestDO::getName, reqVO.getName())
                .eqIfPresent(CgTestDO::getAgeNumber, reqVO.getAgeNumber())
                .betweenIfPresent(CgTestDO::getCreateTime, reqVO.getCreateTime())
//                .inIfPresent(AdminUserDO::getDeptId, deptIds)
                .orderByDesc(CgTestDO::getId));
    }

    default List<CgTestDO> selectListByName(String name) {
        return selectList(new LambdaQueryWrapperX<CgTestDO>().like(CgTestDO::getName, name));
    }

    default List<CgTestDO> selectListByStatus(Integer ageNumber) {
        return selectList(CgTestDO::getAgeNumber, ageNumber);
    }

    default List<CgTestDO> selectListByDeptIds(Collection<Integer> ageNumbers) {
        return selectList(CgTestDO::getAgeNumber, ageNumbers);
    }
}
