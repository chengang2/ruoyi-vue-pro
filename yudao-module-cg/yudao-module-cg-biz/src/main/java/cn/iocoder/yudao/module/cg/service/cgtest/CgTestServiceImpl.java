package cn.iocoder.yudao.module.cg.service.cgtest;

import com.mzt.logapi.context.LogRecordContext;
import cn.iocoder.yudao.framework.common.pojo.PageResult;
import cn.iocoder.yudao.framework.common.util.object.BeanUtils;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgPagePeqVO;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgSaveReqVO;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgUpdateReqVO;
import cn.iocoder.yudao.module.cg.dal.dataobject.cgtest.CgTestDO;
import cn.iocoder.yudao.module.cg.dal.mysql.cgtest.CgTestMapper;
import jakarta.annotation.Resource;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.Collection;
import java.util.List;

import static cn.iocoder.yudao.framework.common.exception.util.ServiceExceptionUtil.exception;
import static cn.iocoder.yudao.module.system.enums.ErrorCodeConstants.USER_NOT_EXISTS;

@Service("cgTestService")
@Slf4j
public class CgTestServiceImpl implements CgTestService {
    @Resource
    private CgTestMapper cgTestMapper;

    @Override
    public Long createUser(CgSaveReqVO createReqVO) {
        CgTestDO user = BeanUtils.toBean(createReqVO, CgTestDO.class);
        log.error("user="+user);
        //user.setAgeNumber(100);
        cgTestMapper.insert(user);
        // 3. 记录操作日志上下文
        LogRecordContext.putVariable("cgtest", user);
        log.error("id:"+user.getId());
        return user.getId();
    }

    @Override
    public void updateUser(CgUpdateReqVO updateReqVO) {
        CgTestDO user = BeanUtils.toBean(updateReqVO, CgTestDO.class);
        log.error("user="+user);
        int i = cgTestMapper.updateById(user);
        log.error("row="+i);
    }

    @Override
    public void deleteUser(Long id) {
        CgTestDO cgTestDO = validateUserExists(id);
        int i = cgTestMapper.deleteById(id);
        log.error("row=="+i);
        LogRecordContext.putVariable("cgtest", cgTestDO);
    }
    @Override
    public void deleteUsers(List<Long> ids) {
        int i = cgTestMapper.deleteBatchIds(ids);
        log.error("row=="+i);
    }

    @Override
    public CgTestDO getUserByUsername(String username) {
        return null;
    }

    @Override
    public CgTestDO getUser(Long id) {
        CgTestDO user = cgTestMapper.selectById(id);
        return user;
    }

    @Override
    public PageResult<CgTestDO> getUserPage(CgPagePeqVO reqVO) {
        PageResult<CgTestDO> cgTestDOPageResult = cgTestMapper.selectPage(reqVO);
        return cgTestDOPageResult;
    }

    @Override
    public List<CgTestDO> getUserList(Collection<Long> ids) {
        return null;
    }

    @Override
    public List<CgTestDO> getUserListByStatus(Integer ageNumber) {
        return null;
    }

    CgTestDO validateUserExists(Long id) {
        if (id == null) {
            return null;
        }
        CgTestDO user = cgTestMapper.selectById(id);
        if (user == null) {
            throw exception(USER_NOT_EXISTS);
        }
        return user;
    }
}
