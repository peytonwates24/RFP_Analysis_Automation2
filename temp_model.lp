\* Sourcing_with_MultiTier_Rebates_Discounts *\
Minimize
OBJ: S_A + S_B + S_C - rebate_A - rebate_B - rebate_C
Subject To
BaseSpend_A: S0_A - 64 x_A_10 - 70 x_A_2 - 55 x_A_3 - 23 x_A_4 - 54 x_A_5
 - 42 x_A_6 - 23 x_A_7 - 75 x_A_8 - 97 x_A_9 = 0
BaseSpend_B: S0_B - 13 x_B_10 - 70 x_B_2 - 65 x_B_3 - 75 x_B_4 - 34 x_B_5
 - 24 x_B_6 - 85 x_B_7 - 13 x_B_8 - 56 x_B_9 = 0
BaseSpend_C: S0_C - 15 x_C_10 - 75 x_C_2 - 60 x_C_3 - 24 x_C_4 - 44 x_C_6
 - 42 x_C_7 - 24 x_C_8 - 32 x_C_9 = 0
Capacity_B_Bid_ID_1: x_B_1 <= 100000000
Capacity_B_Bid_ID_10: x_B_10 <= 100000000
Capacity_B_Bid_ID_2: x_B_2 <= 100000000
Capacity_B_Bid_ID_3: x_B_3 <= 100000000
Capacity_B_Bid_ID_4: x_B_4 <= 100000000
Capacity_B_Bid_ID_5: x_B_5 <= 100000000
Capacity_B_Bid_ID_6: x_B_6 <= 100000000
Capacity_B_Bid_ID_7: x_B_7 <= 100000000
Capacity_B_Bid_ID_8: x_B_8 <= 100000000
Capacity_B_Bid_ID_9: x_B_9 <= 100000000
Capacity_C_Bid_ID_1: x_C_1 <= 100000000
Capacity_C_Bid_ID_10: x_C_10 <= 100000000
Capacity_C_Bid_ID_2: x_C_2 <= 100000000
Capacity_C_Bid_ID_3: x_C_3 <= 100000000
Capacity_C_Bid_ID_4: x_C_4 <= 100000000
Capacity_C_Bid_ID_5: x_C_5 <= 100000000
Capacity_C_Bid_ID_6: x_C_6 <= 100000000
Capacity_C_Bid_ID_7: x_C_7 <= 100000000
Capacity_C_Bid_ID_8: x_C_8 <= 100000000
Capacity_C_Bid_ID_9: x_C_9 <= 100000000
Demand_1: x_A_1 + x_B_1 + x_C_1 = 0
Demand_10: x_A_10 + x_B_10 + x_C_10 = 13
Demand_2: x_A_2 + x_B_2 + x_C_2 = 9000
Demand_3: x_A_3 + x_B_3 + x_C_3 = 600
Demand_4: x_A_4 + x_B_4 + x_C_4 = 5670
Demand_5: x_A_5 + x_B_5 + x_C_5 = 45
Demand_6: x_A_6 + x_B_6 + x_C_6 = 242
Demand_7: x_A_7 + x_B_7 + x_C_7 = 664
Demand_8: x_A_8 + x_B_8 + x_C_8 = 24
Demand_9: x_A_9 + x_B_9 + x_C_9 = 232
DiscountTierLower_A_0: d_A - 19400000000 z_discount_A_0 >= -19400000000
DiscountTierLower_A_1: - 0.01 S0_A + d_A - 19400000000 z_discount_A_1
 >= -19400000000
DiscountTierLower_B_0: d_B - 97000000000 z_discount_B_0 >= -97000000000
DiscountTierLower_B_1: - 0.03 S0_B + d_B - 97000000000 z_discount_B_1
 >= -97000000000
DiscountTierLower_C_0: d_C - 97000000000 z_discount_C_0 >= -97000000000
DiscountTierLower_C_1: - 0.04 S0_C + d_C - 97000000000 z_discount_C_1
 >= -97000000000
DiscountTierMax_A_0: 19400000000 z_discount_A_0 <= 19400001000
DiscountTierMax_B_0: 97000000000 z_discount_B_0 <= 97000000500
DiscountTierMax_C_0: 97000000000 z_discount_C_0 <= 97000000500
_dummy: __dummy = 0
DiscountTierMin_A_0: __dummy >= 0
DiscountTierMin_A_1: x_A_1 + x_A_10 + x_A_3 + x_A_4 + x_A_8 + x_A_9
 - 1000 z_discount_A_1 >= 0
DiscountTierMin_B_0: __dummy >= 0
DiscountTierMin_B_1: x_B_2 + x_B_5 + x_B_6 + x_B_7 - 500 z_discount_B_1 >= 0
DiscountTierMin_C_0: __dummy >= 0
DiscountTierMin_C_1: x_C_1 + x_C_10 + x_C_3 + x_C_4 + x_C_8 + x_C_9
 - 500 z_discount_C_1 >= 0
DiscountTierSelect_A: z_discount_A_0 + z_discount_A_1 = 1
DiscountTierSelect_B: z_discount_B_0 + z_discount_B_1 = 1
DiscountTierSelect_C: z_discount_C_0 + z_discount_C_1 = 1
DiscountTierUpper_A_0: d_A + 19400000000 z_discount_A_0 <= 19400000000
DiscountTierUpper_A_1: - 0.01 S0_A + d_A + 19400000000 z_discount_A_1
 <= 19400000000
DiscountTierUpper_B_0: d_B + 97000000000 z_discount_B_0 <= 97000000000
DiscountTierUpper_B_1: - 0.03 S0_B + d_B + 97000000000 z_discount_B_1
 <= 97000000000
DiscountTierUpper_C_0: d_C + 97000000000 z_discount_C_0 <= 97000000000
DiscountTierUpper_C_1: - 0.04 S0_C + d_C + 97000000000 z_discount_C_1
 <= 97000000000
EffectiveSpend_A: - S0_A + S_A + d_A = 0
EffectiveSpend_B: - S0_B + S_B + d_B = 0
EffectiveSpend_C: - S0_C + S_C + d_C = 0
NonBid_A_1: x_A_1 = 0
NonBid_B_1: x_B_1 = 0
NonBid_C_1: x_C_1 = 0
NonBid_C_5: x_C_5 = 0
RebateTierLower_A_0: rebate_A - 19400000000 y_rebate_A_0 >= -19400000000
RebateTierLower_A_1: - 0.1 S_A + rebate_A - 19400000000 y_rebate_A_1
 >= -19400000000
RebateTierLower_B_0: rebate_B - 97000000000 y_rebate_B_0 >= -97000000000
RebateTierLower_B_1: - 0.05 S_B + rebate_B - 97000000000 y_rebate_B_1
 >= -97000000000
RebateTierLower_C_0: rebate_C - 97000000000 y_rebate_C_0 >= -97000000000
RebateTierLower_C_1: - 0.07 S_C + rebate_C - 97000000000 y_rebate_C_1
 >= -97000000000
RebateTierMax_A_0: 19400000000 y_rebate_A_0 <= 19400000500
RebateTierMax_B_0: 97000000000 y_rebate_B_0 <= 97000000500
RebateTierMax_C_0: 97000000000 y_rebate_C_0 <= 97000000700
RebateTierMin_A_0: __dummy >= 0
RebateTierMin_A_1: - 500 y_rebate_A_1 >= 0
RebateTierMin_B_0: __dummy >= 0
RebateTierMin_B_1: x_B_2 + x_B_5 + x_B_6 + x_B_7 - 500 y_rebate_B_1 >= 0
RebateTierMin_C_0: __dummy >= 0
RebateTierMin_C_1: x_C_1 + x_C_10 + x_C_3 + x_C_4 + x_C_8 + x_C_9
 - 700 y_rebate_C_1 >= 0
RebateTierSelect_A: y_rebate_A_0 + y_rebate_A_1 = 1
RebateTierSelect_B: y_rebate_B_0 + y_rebate_B_1 = 1
RebateTierSelect_C: y_rebate_C_0 + y_rebate_C_1 = 1
RebateTierUpper_A_0: rebate_A + 19400000000 y_rebate_A_0 <= 19400000000
RebateTierUpper_A_1: - 0.1 S_A + rebate_A + 19400000000 y_rebate_A_1
 <= 19400000000
RebateTierUpper_B_0: rebate_B + 97000000000 y_rebate_B_0 <= 97000000000
RebateTierUpper_B_1: - 0.05 S_B + rebate_B + 97000000000 y_rebate_B_1
 <= 97000000000
RebateTierUpper_C_0: rebate_C + 97000000000 y_rebate_C_0 <= 97000000000
RebateTierUpper_C_1: - 0.07 S_C + rebate_C + 97000000000 y_rebate_C_1
 <= 97000000000
Transition_lb_1: - 0.001 T_bid_1 + x_B_1 + x_C_1 >= 0
Transition_lb_10: - 0.001 T_bid_10 + x_A_10 + x_B_10 >= 0
Transition_lb_2: - 0.001 T_bid_2 + x_A_2 + x_C_2 >= 0
Transition_lb_3: - 0.001 T_bid_3 + x_A_3 + x_B_3 >= 0
Transition_lb_4: - 0.001 T_bid_4 + x_A_4 + x_B_4 >= 0
Transition_lb_5: - 0.001 T_bid_5 + x_A_5 + x_B_5 >= 0
Transition_lb_6: - 0.001 T_bid_6 + x_A_6 + x_B_6 >= 0
Transition_lb_7: - 0.001 T_bid_7 + x_A_7 + x_B_7 >= 0
Transition_lb_8: - 0.001 T_bid_8 + x_A_8 + x_B_8 >= 0
Transition_lb_9: - 0.001 T_bid_9 + x_A_9 + x_B_9 >= 0
Transition_ub_1: x_B_1 + x_C_1 <= 0
Transition_ub_10: - 13 T_bid_10 + x_A_10 + x_B_10 <= 0
Transition_ub_2: - 9000 T_bid_2 + x_A_2 + x_C_2 <= 0
Transition_ub_3: - 600 T_bid_3 + x_A_3 + x_B_3 <= 0
Transition_ub_4: - 5670 T_bid_4 + x_A_4 + x_B_4 <= 0
Transition_ub_5: - 45 T_bid_5 + x_A_5 + x_B_5 <= 0
Transition_ub_6: - 242 T_bid_6 + x_A_6 + x_B_6 <= 0
Transition_ub_7: - 664 T_bid_7 + x_A_7 + x_B_7 <= 0
Transition_ub_8: - 24 T_bid_8 + x_A_8 + x_B_8 <= 0
Transition_ub_9: - 232 T_bid_9 + x_A_9 + x_B_9 <= 0
Volume_A: V_A - x_A_1 - x_A_10 - x_A_2 - x_A_3 - x_A_4 - x_A_5 - x_A_6 - x_A_7
 - x_A_8 - x_A_9 = 0
Volume_B: V_B - x_B_1 - x_B_10 - x_B_2 - x_B_3 - x_B_4 - x_B_5 - x_B_6 - x_B_7
 - x_B_8 - x_B_9 = 0
Volume_C: V_C - x_C_1 - x_C_10 - x_C_2 - x_C_3 - x_C_4 - x_C_5 - x_C_6 - x_C_7
 - x_C_8 - x_C_9 = 0
Binaries
T_bid_1
T_bid_10
T_bid_2
T_bid_3
T_bid_4
T_bid_5
T_bid_6
T_bid_7
T_bid_8
T_bid_9
y_rebate_A_0
y_rebate_A_1
y_rebate_B_0
y_rebate_B_1
y_rebate_C_0
y_rebate_C_1
z_discount_A_0
z_discount_A_1
z_discount_B_0
z_discount_B_1
z_discount_C_0
z_discount_C_1
End
