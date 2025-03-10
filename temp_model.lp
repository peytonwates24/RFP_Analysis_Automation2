\* Sourcing_with_MultiTier_Rebates_Discounts *\
Minimize
OBJ: S_A + S_B + S_C - rebate_A - rebate_B - rebate_C
Subject To
BaseSpend_A: S0_A - 50 x_A_1 - 64 x_A_10 - 70 x_A_2 - 55 x_A_3 - 23 x_A_4
 - 54 x_A_5 - 42 x_A_6 - 23 x_A_7 - 75 x_A_8 - 97 x_A_9 = 0
BaseSpend_B: S0_B - 60 x_B_1 - 13 x_B_10 - 80 x_B_2 - 65 x_B_3 - 75 x_B_4
 - 34 x_B_5 - 24 x_B_6 - 53 x_B_7 - 13 x_B_8 - 56 x_B_9 = 0
BaseSpend_C: S0_C - 55 x_C_1 - 75 x_C_10 - 75 x_C_2 - 60 x_C_3 - 24 x_C_4
 - 24 x_C_5 - 64 x_C_6 - 86 x_C_7 - 24 x_C_8 - 13 x_C_9 = 0
Capacity_A_BusinessUnit_A: x_A_1 + x_A_10 + x_A_3 + x_A_4 + x_A_5 + x_A_9
 <= 500
Capacity_A_BusinessUnit_B: x_A_2 + x_A_6 + x_A_7 + x_A_8 <= 1000
Capacity_B_Bid_ID_1: x_B_1 <= 4000
Capacity_B_Bid_ID_10: x_B_10 <= 5000
Capacity_B_Bid_ID_2: x_B_2 <= 8000
Capacity_B_Bid_ID_3: x_B_3 <= 6000
Capacity_B_Bid_ID_4: x_B_4 <= 4000
Capacity_B_Bid_ID_5: x_B_5 <= 8000
Capacity_B_Bid_ID_6: x_B_6 <= 6000
Capacity_B_Bid_ID_7: x_B_7 <= 4000
Capacity_B_Bid_ID_8: x_B_8 <= 8000
Capacity_B_Bid_ID_9: x_B_9 <= 6000
Capacity_C_Bid_ID_1: x_C_1 <= 3000
Capacity_C_Bid_ID_10: x_C_10 <= 7000
Capacity_C_Bid_ID_2: x_C_2 <= 5000
Capacity_C_Bid_ID_3: x_C_3 <= 7000
Capacity_C_Bid_ID_4: x_C_4 <= 3000
Capacity_C_Bid_ID_5: x_C_5 <= 5000
Capacity_C_Bid_ID_6: x_C_6 <= 7000
Capacity_C_Bid_ID_7: x_C_7 <= 3000
Capacity_C_Bid_ID_8: x_C_8 <= 5000
Capacity_C_Bid_ID_9: x_C_9 <= 7000
Demand_1: x_A_1 + x_B_1 + x_C_1 = 700
Demand_10: x_A_10 + x_B_10 + x_C_10 = 13
Demand_2: x_A_2 + x_B_2 + x_C_2 = 9000
Demand_3: x_A_3 + x_B_3 + x_C_3 = 600
Demand_4: x_A_4 + x_B_4 + x_C_4 = 5670
Demand_5: x_A_5 + x_B_5 + x_C_5 = 45
Demand_6: x_A_6 + x_B_6 + x_C_6 = 242
Demand_7: x_A_7 + x_B_7 + x_C_7 = 664
Demand_8: x_A_8 + x_B_8 + x_C_8 = 24
Demand_9: x_A_9 + x_B_9 + x_C_9 = 232
DiscountTierLower_A_0: d_A - 145500 z_discount_A_0 >= -145500
DiscountTierLower_A_1: - 0.01 S0_A + d_A - 145500 z_discount_A_1 >= -145500
DiscountTierLower_B_0: d_B - 5723000 z_discount_B_0 >= -5723000
DiscountTierLower_B_1: - 0.03 S0_B + d_B - 5723000 z_discount_B_1 >= -5723000
DiscountTierLower_C_0: d_C - 5044000 z_discount_C_0 >= -5044000
DiscountTierLower_C_1: - 0.04 S0_C + d_C - 5044000 z_discount_C_1 >= -5044000
DiscountTierMax_A_0: 145500 z_discount_A_0 <= 146500
DiscountTierMax_B_0: 5723000 z_discount_B_0 <= 5723500
DiscountTierMax_C_0: 5044000 z_discount_C_0 <= 5044500
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
DiscountTierUpper_A_0: d_A + 145500 z_discount_A_0 <= 145500
DiscountTierUpper_A_1: - 0.01 S0_A + d_A + 145500 z_discount_A_1 <= 145500
DiscountTierUpper_B_0: d_B + 5723000 z_discount_B_0 <= 5723000
DiscountTierUpper_B_1: - 0.03 S0_B + d_B + 5723000 z_discount_B_1 <= 5723000
DiscountTierUpper_C_0: d_C + 5044000 z_discount_C_0 <= 5044000
DiscountTierUpper_C_1: - 0.04 S0_C + d_C + 5044000 z_discount_C_1 <= 5044000
EffectiveSpend_A: - S0_A + S_A + d_A = 0
EffectiveSpend_B: - S0_B + S_B + d_B = 0
EffectiveSpend_C: - S0_C + S_C + d_C = 0
RebateTierLower_A_0: rebate_A - 145500 y_rebate_A_0 >= -145500
RebateTierLower_A_1: - 0.1 S_A + rebate_A - 145500 y_rebate_A_1 >= -145500
RebateTierLower_B_0: rebate_B - 5723000 y_rebate_B_0 >= -5723000
RebateTierLower_B_1: - 0.05 S_B + rebate_B - 5723000 y_rebate_B_1 >= -5723000
RebateTierLower_C_0: rebate_C - 5044000 y_rebate_C_0 >= -5044000
RebateTierLower_C_1: - 0.07 S_C + rebate_C - 5044000 y_rebate_C_1 >= -5044000
RebateTierMax_A_0: 145500 y_rebate_A_0 <= 146000
RebateTierMax_B_0: 5723000 y_rebate_B_0 <= 5723500
RebateTierMax_C_0: 5044000 y_rebate_C_0 <= 5044700
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
RebateTierUpper_A_0: rebate_A + 145500 y_rebate_A_0 <= 145500
RebateTierUpper_A_1: - 0.1 S_A + rebate_A + 145500 y_rebate_A_1 <= 145500
RebateTierUpper_B_0: rebate_B + 5723000 y_rebate_B_0 <= 5723000
RebateTierUpper_B_1: - 0.05 S_B + rebate_B + 5723000 y_rebate_B_1 <= 5723000
RebateTierUpper_C_0: rebate_C + 5044000 y_rebate_C_0 <= 5044000
RebateTierUpper_C_1: - 0.07 S_C + rebate_C + 5044000 y_rebate_C_1 <= 5044000
RuleSingleSupplier_0_1: w_0_1_A + w_0_1_B + w_0_1_C = 1
RuleSingleSupplier_1_2: w_1_2_A + w_1_2_B + w_1_2_C = 1
RuleSingleSupplier_2_10: w_2_10_A + w_2_10_B + w_2_10_C = 1
RuleSingleSupplier_3_3: w_3_3_A + w_3_3_B + w_3_3_C = 1
RuleSingleSupplier_4_4: w_4_4_A + w_4_4_B + w_4_4_C = 1
RuleSingleSupplier_5_5: w_5_5_A + w_5_5_B + w_5_5_C = 1
RuleSingleSupplier_6_6: w_6_6_A + w_6_6_B + w_6_6_C = 1
RuleSingleSupplier_7_7: w_7_7_A + w_7_7_B + w_7_7_C = 1
RuleSingleSupplier_8_8: w_8_8_A + w_8_8_B + w_8_8_C = 1
RuleSingleSupplier_9_9: w_9_9_A + w_9_9_B + w_9_9_C = 1
RuleSupplierIndicatorLB_0_1_A: - 0.001 w_0_1_A + x_A_1 >= 0
RuleSupplierIndicatorLB_0_1_B: - 0.001 w_0_1_B + x_B_1 >= 0
RuleSupplierIndicatorLB_0_1_C: - 0.001 w_0_1_C + x_C_1 >= 0
RuleSupplierIndicatorLB_1_2_A: - 0.001 w_1_2_A + x_A_2 >= 0
RuleSupplierIndicatorLB_1_2_B: - 0.001 w_1_2_B + x_B_2 >= 0
RuleSupplierIndicatorLB_1_2_C: - 0.001 w_1_2_C + x_C_2 >= 0
RuleSupplierIndicatorLB_2_10_A: - 0.001 w_2_10_A + x_A_10 >= 0
RuleSupplierIndicatorLB_2_10_B: - 0.001 w_2_10_B + x_B_10 >= 0
RuleSupplierIndicatorLB_2_10_C: - 0.001 w_2_10_C + x_C_10 >= 0
RuleSupplierIndicatorLB_3_3_A: - 0.001 w_3_3_A + x_A_3 >= 0
RuleSupplierIndicatorLB_3_3_B: - 0.001 w_3_3_B + x_B_3 >= 0
RuleSupplierIndicatorLB_3_3_C: - 0.001 w_3_3_C + x_C_3 >= 0
RuleSupplierIndicatorLB_4_4_A: - 0.001 w_4_4_A + x_A_4 >= 0
RuleSupplierIndicatorLB_4_4_B: - 0.001 w_4_4_B + x_B_4 >= 0
RuleSupplierIndicatorLB_4_4_C: - 0.001 w_4_4_C + x_C_4 >= 0
RuleSupplierIndicatorLB_5_5_A: - 0.001 w_5_5_A + x_A_5 >= 0
RuleSupplierIndicatorLB_5_5_B: - 0.001 w_5_5_B + x_B_5 >= 0
RuleSupplierIndicatorLB_5_5_C: - 0.001 w_5_5_C + x_C_5 >= 0
RuleSupplierIndicatorLB_6_6_A: - 0.001 w_6_6_A + x_A_6 >= 0
RuleSupplierIndicatorLB_6_6_B: - 0.001 w_6_6_B + x_B_6 >= 0
RuleSupplierIndicatorLB_6_6_C: - 0.001 w_6_6_C + x_C_6 >= 0
RuleSupplierIndicatorLB_7_7_A: - 0.001 w_7_7_A + x_A_7 >= 0
RuleSupplierIndicatorLB_7_7_B: - 0.001 w_7_7_B + x_B_7 >= 0
RuleSupplierIndicatorLB_7_7_C: - 0.001 w_7_7_C + x_C_7 >= 0
RuleSupplierIndicatorLB_8_8_A: - 0.001 w_8_8_A + x_A_8 >= 0
RuleSupplierIndicatorLB_8_8_B: - 0.001 w_8_8_B + x_B_8 >= 0
RuleSupplierIndicatorLB_8_8_C: - 0.001 w_8_8_C + x_C_8 >= 0
RuleSupplierIndicatorLB_9_9_A: - 0.001 w_9_9_A + x_A_9 >= 0
RuleSupplierIndicatorLB_9_9_B: - 0.001 w_9_9_B + x_B_9 >= 0
RuleSupplierIndicatorLB_9_9_C: - 0.001 w_9_9_C + x_C_9 >= 0
RuleSupplierIndicator_0_1_A: - 1000000000 w_0_1_A + x_A_1 <= 0
RuleSupplierIndicator_0_1_B: - 1000000000 w_0_1_B + x_B_1 <= 0
RuleSupplierIndicator_0_1_C: - 1000000000 w_0_1_C + x_C_1 <= 0
RuleSupplierIndicator_1_2_A: - 1000000000 w_1_2_A + x_A_2 <= 0
RuleSupplierIndicator_1_2_B: - 1000000000 w_1_2_B + x_B_2 <= 0
RuleSupplierIndicator_1_2_C: - 1000000000 w_1_2_C + x_C_2 <= 0
RuleSupplierIndicator_2_10_A: - 1000000000 w_2_10_A + x_A_10 <= 0
RuleSupplierIndicator_2_10_B: - 1000000000 w_2_10_B + x_B_10 <= 0
RuleSupplierIndicator_2_10_C: - 1000000000 w_2_10_C + x_C_10 <= 0
RuleSupplierIndicator_3_3_A: - 1000000000 w_3_3_A + x_A_3 <= 0
RuleSupplierIndicator_3_3_B: - 1000000000 w_3_3_B + x_B_3 <= 0
RuleSupplierIndicator_3_3_C: - 1000000000 w_3_3_C + x_C_3 <= 0
RuleSupplierIndicator_4_4_A: - 1000000000 w_4_4_A + x_A_4 <= 0
RuleSupplierIndicator_4_4_B: - 1000000000 w_4_4_B + x_B_4 <= 0
RuleSupplierIndicator_4_4_C: - 1000000000 w_4_4_C + x_C_4 <= 0
RuleSupplierIndicator_5_5_A: - 1000000000 w_5_5_A + x_A_5 <= 0
RuleSupplierIndicator_5_5_B: - 1000000000 w_5_5_B + x_B_5 <= 0
RuleSupplierIndicator_5_5_C: - 1000000000 w_5_5_C + x_C_5 <= 0
RuleSupplierIndicator_6_6_A: - 1000000000 w_6_6_A + x_A_6 <= 0
RuleSupplierIndicator_6_6_B: - 1000000000 w_6_6_B + x_B_6 <= 0
RuleSupplierIndicator_6_6_C: - 1000000000 w_6_6_C + x_C_6 <= 0
RuleSupplierIndicator_7_7_A: - 1000000000 w_7_7_A + x_A_7 <= 0
RuleSupplierIndicator_7_7_B: - 1000000000 w_7_7_B + x_B_7 <= 0
RuleSupplierIndicator_7_7_C: - 1000000000 w_7_7_C + x_C_7 <= 0
RuleSupplierIndicator_8_8_A: - 1000000000 w_8_8_A + x_A_8 <= 0
RuleSupplierIndicator_8_8_B: - 1000000000 w_8_8_B + x_B_8 <= 0
RuleSupplierIndicator_8_8_C: - 1000000000 w_8_8_C + x_C_8 <= 0
RuleSupplierIndicator_9_9_A: - 1000000000 w_9_9_A + x_A_9 <= 0
RuleSupplierIndicator_9_9_B: - 1000000000 w_9_9_B + x_B_9 <= 0
RuleSupplierIndicator_9_9_C: - 1000000000 w_9_9_C + x_C_9 <= 0
Transition_10_A: - 13 T_10_A + x_A_10 <= 0
Transition_10_B: - 13 T_10_B + x_B_10 <= 0
Transition_1_B: - 700 T_1_B + x_B_1 <= 0
Transition_1_C: - 700 T_1_C + x_C_1 <= 0
Transition_2_A: - 9000 T_2_A + x_A_2 <= 0
Transition_2_C: - 9000 T_2_C + x_C_2 <= 0
Transition_3_A: - 600 T_3_A + x_A_3 <= 0
Transition_3_B: - 600 T_3_B + x_B_3 <= 0
Transition_4_A: - 5670 T_4_A + x_A_4 <= 0
Transition_4_B: - 5670 T_4_B + x_B_4 <= 0
Transition_5_A: - 45 T_5_A + x_A_5 <= 0
Transition_5_B: - 45 T_5_B + x_B_5 <= 0
Transition_6_A: - 242 T_6_A + x_A_6 <= 0
Transition_6_B: - 242 T_6_B + x_B_6 <= 0
Transition_7_A: - 664 T_7_A + x_A_7 <= 0
Transition_7_B: - 664 T_7_B + x_B_7 <= 0
Transition_8_A: - 24 T_8_A + x_A_8 <= 0
Transition_8_B: - 24 T_8_B + x_B_8 <= 0
Transition_9_A: - 232 T_9_A + x_A_9 <= 0
Transition_9_B: - 232 T_9_B + x_B_9 <= 0
Volume_A: V_A - x_A_1 - x_A_10 - x_A_2 - x_A_3 - x_A_4 - x_A_5 - x_A_6 - x_A_7
 - x_A_8 - x_A_9 = 0
Volume_B: V_B - x_B_1 - x_B_10 - x_B_2 - x_B_3 - x_B_4 - x_B_5 - x_B_6 - x_B_7
 - x_B_8 - x_B_9 = 0
Volume_C: V_C - x_C_1 - x_C_10 - x_C_2 - x_C_3 - x_C_4 - x_C_5 - x_C_6 - x_C_7
 - x_C_8 - x_C_9 = 0
Bounds
 __dummy = 0
Binaries
T_10_A
T_10_B
T_1_B
T_1_C
T_2_A
T_2_C
T_3_A
T_3_B
T_4_A
T_4_B
T_5_A
T_5_B
T_6_A
T_6_B
T_7_A
T_7_B
T_8_A
T_8_B
T_9_A
T_9_B
w_0_1_A
w_0_1_B
w_0_1_C
w_1_2_A
w_1_2_B
w_1_2_C
w_2_10_A
w_2_10_B
w_2_10_C
w_3_3_A
w_3_3_B
w_3_3_C
w_4_4_A
w_4_4_B
w_4_4_C
w_5_5_A
w_5_5_B
w_5_5_C
w_6_6_A
w_6_6_B
w_6_6_C
w_7_7_A
w_7_7_B
w_7_7_C
w_8_8_A
w_8_8_B
w_8_8_C
w_9_9_A
w_9_9_B
w_9_9_C
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
