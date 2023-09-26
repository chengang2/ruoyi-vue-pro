package cn.iocoder.yudao.module.trade.service.price.calculator;

import cn.hutool.core.util.BooleanUtil;
import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import cn.iocoder.yudao.module.member.api.point.MemberPointApi;
import cn.iocoder.yudao.module.member.api.point.dto.MemberPointConfigRespDTO;
import cn.iocoder.yudao.module.member.api.user.MemberUserApi;
import cn.iocoder.yudao.module.member.api.user.dto.MemberUserRespDTO;
import cn.iocoder.yudao.module.promotion.enums.common.PromotionTypeEnum;
import cn.iocoder.yudao.module.trade.service.price.bo.TradePriceCalculateReqBO;
import cn.iocoder.yudao.module.trade.service.price.bo.TradePriceCalculateRespBO;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.annotation.Order;
import org.springframework.stereotype.Component;

import javax.annotation.Resource;
import java.math.RoundingMode;
import java.util.List;

import static cn.iocoder.yudao.framework.common.util.collection.CollectionUtils.filterList;

// TODO @疯狂：搞个单测，嘿嘿；
/**
 * 使用积分的 {@link TradePriceCalculator} 实现类
 *
 * @author owen
 */
@Component
@Order(TradePriceCalculator.ORDER_POINT_USE)
@Slf4j
public class TradePointUsePriceCalculator implements TradePriceCalculator {

    @Resource
    private MemberPointApi memberPointApi;
    @Resource
    private MemberUserApi memberUserApi;

    @Override
    public void calculate(TradePriceCalculateReqBO param, TradePriceCalculateRespBO result) {
        // 1.1 校验是否使用积分
        if (!BooleanUtil.isTrue(param.getPointStatus())) {
            result.setUsePoint(0);
            return;
        }
        // 1.2 校验积分抵扣是否开启
        MemberPointConfigRespDTO config = memberPointApi.getConfig();
        if (!checkDeductPointEnable(config)) {
            return;
        }
        // 1.3 校验用户积分余额
        MemberUserRespDTO user = memberUserApi.getUser(param.getUserId());
        if (user.getPoint() == null || user.getPoint() <= 0) {
            return;
        }

        // 2.1 计算积分优惠金额
        // TODO @疯狂：如果计算出来，优惠金额为 0，那是不是不用执行后续逻辑哈
        int pointPrice = calculatePointPrice(config, user.getPoint(), result);
        // 2.2 计算分摊的积分、抵扣金额
        List<TradePriceCalculateRespBO.OrderItem> orderItems = filterList(result.getItems(), TradePriceCalculateRespBO.OrderItem::getSelected);
        List<Integer> dividePointPrices = TradePriceCalculatorHelper.dividePrice(orderItems, pointPrice);
        List<Integer> divideUsePoints = TradePriceCalculatorHelper.dividePrice(orderItems, result.getUsePoint());

        // 3.1 记录优惠明细
        TradePriceCalculatorHelper.addPromotion(result, orderItems,
                param.getUserId(), "积分抵扣", PromotionTypeEnum.POINT.getType(),
                StrUtil.format("积分抵扣：省 {} 元", TradePriceCalculatorHelper.formatPrice(pointPrice)),
                dividePointPrices);
        // 3.2 更新 SKU 优惠金额
        for (int i = 0; i < orderItems.size(); i++) {
            TradePriceCalculateRespBO.OrderItem orderItem = orderItems.get(i);
            orderItem.setPointPrice(dividePointPrices.get(i));
            orderItem.setUsePoint(divideUsePoints.get(i));
            TradePriceCalculatorHelper.recountPayPrice(orderItem);
        }
        TradePriceCalculatorHelper.recountAllPrice(result);
    }

    // TODO @疯狂：这个最好是 is 开头；因为 check 或者 validator，更多失败，会抛出异常；
    private boolean checkDeductPointEnable(MemberPointConfigRespDTO config) {
        // TODO @疯狂：这个要不直接写成 return config != null && config .... 多行这样一个形式；
        if (config == null) {
            return false;
        }
        if (!BooleanUtil.isTrue(config.getTradeDeductEnable())) {
            return false;
        }

        // 有没有配置：1 积分抵扣多少分
        return config.getTradeDeductUnitPrice() != null && config.getTradeDeductUnitPrice() > 0;
    }

    private Integer calculatePointPrice(MemberPointConfigRespDTO config, Integer usePoint, TradePriceCalculateRespBO result) {
        // 每个订单最多可以使用的积分数量
        if (config.getTradeDeductMaxPrice() != null && config.getTradeDeductMaxPrice() > 0) {
            usePoint = Math.min(usePoint, config.getTradeDeductMaxPrice());
        }
        // 积分优惠金额（分）
        int pointPrice = usePoint * config.getTradeDeductUnitPrice();
        // 0 元购!!!：用户积分比较多时，积分可以抵扣的金额要大于支付金额，这时需要根据支付金额反推使用多少积分
        if (result.getPrice().getPayPrice() < pointPrice) {
            pointPrice = result.getPrice().getPayPrice();
            // 反推需要扣除的积分
            usePoint = NumberUtil.toBigDecimal(pointPrice)
                    .divide(NumberUtil.toBigDecimal(config.getTradeDeductUnitPrice()), 0, RoundingMode.HALF_UP)
                    .intValue();
        }
        // 记录使用的积分
        result.setUsePoint(usePoint);
        return pointPrice;
    }

}