\* Sourcing_with_MultiTier_Rebates_Discounts *\
Minimize
OBJ: S_BASF + S_Covestro + S_Dow + S_Huntsman + S_Wanhua - rebate_BASF
 - rebate_Covestro - rebate_Dow - rebate_Huntsman - rebate_Wanhua
Subject To
ActiveLink_BASF_1: x_BASF_1 - 1000000000 z_BASF <= 0
ActiveLink_BASF_2: x_BASF_2 - 1000000000 z_BASF <= 0
ActiveLink_BASF_3: x_BASF_3 - 1000000000 z_BASF <= 0
ActiveLink_BASF_4: x_BASF_4 - 1000000000 z_BASF <= 0
ActiveLink_BASF_5: x_BASF_5 - 1000000000 z_BASF <= 0
ActiveLink_Covestro_1: x_Covestro_1 - 1000000000 z_Covestro <= 0
ActiveLink_Covestro_2: x_Covestro_2 - 1000000000 z_Covestro <= 0
ActiveLink_Covestro_3: x_Covestro_3 - 1000000000 z_Covestro <= 0
ActiveLink_Covestro_4: x_Covestro_4 - 1000000000 z_Covestro <= 0
ActiveLink_Covestro_5: x_Covestro_5 - 1000000000 z_Covestro <= 0
ActiveLink_Dow_1: x_Dow_1 - 1000000000 z_Dow <= 0
ActiveLink_Dow_2: x_Dow_2 - 1000000000 z_Dow <= 0
ActiveLink_Dow_3: x_Dow_3 - 1000000000 z_Dow <= 0
ActiveLink_Dow_4: x_Dow_4 - 1000000000 z_Dow <= 0
ActiveLink_Dow_5: x_Dow_5 - 1000000000 z_Dow <= 0
ActiveLink_Huntsman_1: x_Huntsman_1 - 1000000000 z_Huntsman <= 0
ActiveLink_Huntsman_2: x_Huntsman_2 - 1000000000 z_Huntsman <= 0
ActiveLink_Huntsman_3: x_Huntsman_3 - 1000000000 z_Huntsman <= 0
ActiveLink_Huntsman_4: x_Huntsman_4 - 1000000000 z_Huntsman <= 0
ActiveLink_Huntsman_5: x_Huntsman_5 - 1000000000 z_Huntsman <= 0
ActiveLink_Wanhua_1: x_Wanhua_1 - 1000000000 z_Wanhua <= 0
ActiveLink_Wanhua_2: x_Wanhua_2 - 1000000000 z_Wanhua <= 0
ActiveLink_Wanhua_3: x_Wanhua_3 - 1000000000 z_Wanhua <= 0
ActiveLink_Wanhua_4: x_Wanhua_4 - 1000000000 z_Wanhua <= 0
ActiveLink_Wanhua_5: x_Wanhua_5 - 1000000000 z_Wanhua <= 0
BaseSpend_BASF: S0_BASF - 0.8762 x_BASF_1 - 0.8862 x_BASF_2 - 0.8762 x_BASF_3
 - 0.9372 x_BASF_4 - 0.8762 x_BASF_5 = 0
BaseSpend_Covestro: S0_Covestro - 0.788 x_Covestro_1 - 0.973 x_Covestro_2
 - 0.863 x_Covestro_3 - 0.983 x_Covestro_4 - 0.813 x_Covestro_5 = 0
BaseSpend_Dow: S0_Dow - 0.753 x_Dow_1 - 0.823 x_Dow_2 - 0.773 x_Dow_3
 - 0.843 x_Dow_4 - 0.743 x_Dow_5 = 0
BaseSpend_Huntsman: S0_Huntsman - 0.83 x_Huntsman_1 - 0.83 x_Huntsman_2
 - 0.83 x_Huntsman_3 - 0.83 x_Huntsman_4 - 0.83 x_Huntsman_5 = 0
BaseSpend_Wanhua: S0_Wanhua - 0.8 x_Wanhua_1 - 0.8 x_Wanhua_2 - 0.8 x_Wanhua_3
 - 0.831 x_Wanhua_4 - 0.8 x_Wanhua_5 = 0
Capacity_BASF_Bid_ID_1: x_BASF_1 <= 10000000000
Capacity_BASF_Bid_ID_2: x_BASF_2 <= 10000000000
Capacity_BASF_Bid_ID_3: x_BASF_3 <= 10000000000
Capacity_BASF_Bid_ID_4: x_BASF_4 <= 10000000000
Capacity_BASF_Bid_ID_5: x_BASF_5 <= 10000000000
Capacity_Covestro_Bid_ID_1: x_Covestro_1 <= 10000000000
Capacity_Covestro_Bid_ID_2: x_Covestro_2 <= 10000000000
Capacity_Covestro_Bid_ID_3: x_Covestro_3 <= 10000000000
Capacity_Covestro_Bid_ID_4: x_Covestro_4 <= 10000000000
Capacity_Covestro_Bid_ID_5: x_Covestro_5 <= 10000000000
Capacity_Dow_Bid_ID_1: x_Dow_1 <= 10000000000
Capacity_Dow_Bid_ID_2: x_Dow_2 <= 10000000000
Capacity_Dow_Bid_ID_3: x_Dow_3 <= 10000000000
Capacity_Dow_Bid_ID_4: x_Dow_4 <= 10000000000
Capacity_Dow_Bid_ID_5: x_Dow_5 <= 10000000000
Capacity_Huntsman_Bid_ID_1: x_Huntsman_1 <= 10000000000
Capacity_Huntsman_Bid_ID_2: x_Huntsman_2 <= 10000000000
Capacity_Huntsman_Bid_ID_3: x_Huntsman_3 <= 10000000000
Capacity_Huntsman_Bid_ID_4: x_Huntsman_4 <= 10000000000
Capacity_Huntsman_Bid_ID_5: x_Huntsman_5 <= 10000000000
Capacity_Wanhua_Bid_ID_1: x_Wanhua_1 <= 10000000000
Capacity_Wanhua_Bid_ID_2: x_Wanhua_2 <= 10000000000
Capacity_Wanhua_Bid_ID_3: x_Wanhua_3 <= 10000000000
Capacity_Wanhua_Bid_ID_4: x_Wanhua_4 <= 10000000000
Capacity_Wanhua_Bid_ID_5: x_Wanhua_5 <= 10000000000
Demand_1: x_BASF_1 + x_Covestro_1 + x_Dow_1 + x_Huntsman_1 + x_Wanhua_1
 = 8860493
Demand_2: x_BASF_2 + x_Covestro_2 + x_Dow_2 + x_Huntsman_2 + x_Wanhua_2
 = 7093293
Demand_3: x_BASF_3 + x_Covestro_3 + x_Dow_3 + x_Huntsman_3 + x_Wanhua_3
 = 7423090
Demand_4: x_BASF_4 + x_Covestro_4 + x_Dow_4 + x_Huntsman_4 + x_Wanhua_4
 = 14146424
Demand_5: x_BASF_5 + x_Covestro_5 + x_Dow_5 + x_Huntsman_5 + x_Wanhua_5
 = 4943948
EffectiveSpend_BASF: - S0_BASF + S_BASF + d_BASF = 0
EffectiveSpend_Covestro: - S0_Covestro + S_Covestro + d_Covestro = 0
EffectiveSpend_Dow: - S0_Dow + S_Dow + d_Dow = 0
EffectiveSpend_Huntsman: - S0_Huntsman + S_Huntsman + d_Huntsman = 0
EffectiveSpend_Wanhua: - S0_Wanhua + S_Wanhua + d_Wanhua = 0
Fix_d_BASF: d_BASF = 0
Fix_d_Covestro: d_Covestro = 0
Fix_d_Dow: d_Dow = 0
Fix_d_Huntsman: d_Huntsman = 0
Fix_d_Wanhua: d_Wanhua = 0
Fix_rebate_Dow: rebate_Dow = 0
Fix_rebate_Wanhua: rebate_Wanhua = 0
MinAward_BASF: x_BASF_1 + x_BASF_2 + x_BASF_3 + x_BASF_4 + x_BASF_5 - z_BASF
 >= 0
MinAward_Covestro: x_Covestro_1 + x_Covestro_2 + x_Covestro_3 + x_Covestro_4
 + x_Covestro_5 - z_Covestro >= 0
MinAward_Dow: x_Dow_1 + x_Dow_2 + x_Dow_3 + x_Dow_4 + x_Dow_5 - z_Dow >= 0
MinAward_Huntsman: x_Huntsman_1 + x_Huntsman_2 + x_Huntsman_3 + x_Huntsman_4
 + x_Huntsman_5 - z_Huntsman >= 0
MinAward_Wanhua: x_Wanhua_1 + x_Wanhua_2 + x_Wanhua_3 + x_Wanhua_4
 + x_Wanhua_5 - z_Wanhua >= 0
RebateTierLower_BASF_0: rebate_BASF - 49150000000 y_rebate_BASF_0
 >= -49150000000
RebateTierLower_BASF_1: rebate_BASF - 49150000000 y_rebate_BASF_1
 >= -49150000000
RebateTierLower_Covestro_0: rebate_Covestro - 49150000000 y_rebate_Covestro_0
 >= -49150000000
RebateTierLower_Covestro_1: rebate_Covestro - 49150000000 y_rebate_Covestro_1
 >= -49150000000
RebateTierLower_Huntsman_0: rebate_Huntsman - 49150000000 y_rebate_Huntsman_0
 >= -49150000000
RebateTierLower_Huntsman_1: rebate_Huntsman - 49150000000 y_rebate_Huntsman_1
 >= -49150000000
RebateTierMax_BASF_0: 49150000000 y_rebate_BASF_0 <= 49150000500
RebateTierMax_Covestro_0: 49150000000 y_rebate_Covestro_0 <= 49150000500
RebateTierMax_Huntsman_0: 49150000000 y_rebate_Huntsman_0 <= 49150000700
_dummy: __dummy = 0
RebateTierMin_BASF_0: __dummy >= 0
RebateTierMin_BASF_1: - 500 y_rebate_BASF_1 >= 0
RebateTierMin_Covestro_0: __dummy >= 0
RebateTierMin_Covestro_1: x_Covestro_1 + x_Covestro_2 + x_Covestro_3
 + x_Covestro_4 + x_Covestro_5 - 500 y_rebate_Covestro_1 >= 0
RebateTierMin_Huntsman_0: __dummy >= 0
RebateTierMin_Huntsman_1: x_Huntsman_1 + x_Huntsman_2 + x_Huntsman_3
 + x_Huntsman_4 + x_Huntsman_5 - 700 y_rebate_Huntsman_1 >= 0
RebateTierSelect_BASF: y_rebate_BASF_0 + y_rebate_BASF_1 = 1
RebateTierSelect_Covestro: y_rebate_Covestro_0 + y_rebate_Covestro_1 = 1
RebateTierSelect_Huntsman: y_rebate_Huntsman_0 + y_rebate_Huntsman_1 = 1
RebateTierUpper_BASF_0: rebate_BASF + 49150000000 y_rebate_BASF_0
 <= 49150000000
RebateTierUpper_BASF_1: rebate_BASF + 49150000000 y_rebate_BASF_1
 <= 49150000000
RebateTierUpper_Covestro_0: rebate_Covestro + 49150000000 y_rebate_Covestro_0
 <= 49150000000
RebateTierUpper_Covestro_1: rebate_Covestro + 49150000000 y_rebate_Covestro_1
 <= 49150000000
RebateTierUpper_Huntsman_0: rebate_Huntsman + 49150000000 y_rebate_Huntsman_0
 <= 49150000000
RebateTierUpper_Huntsman_1: rebate_Huntsman + 49150000000 y_rebate_Huntsman_1
 <= 49150000000
SupplierExclusion_0_1_BASF: x_BASF_1 = 0
SupplierExclusion_0_1_Covestro: x_Covestro_1 = 0
SupplierExclusion_0_1_Huntsman: x_Huntsman_1 = 0
SupplierExclusion_0_1_Wanhua: x_Wanhua_1 = 0
SupplierExclusion_0_2_Covestro: x_Covestro_2 = 0
SupplierExclusion_0_2_Dow: x_Dow_2 = 0
SupplierExclusion_0_2_Huntsman: x_Huntsman_2 = 0
SupplierExclusion_0_2_Wanhua: x_Wanhua_2 = 0
SupplierExclusion_0_3_Covestro: x_Covestro_3 = 0
SupplierExclusion_0_3_Dow: x_Dow_3 = 0
SupplierExclusion_0_3_Huntsman: x_Huntsman_3 = 0
SupplierExclusion_0_3_Wanhua: x_Wanhua_3 = 0
SupplierExclusion_0_4_Covestro: x_Covestro_4 = 0
SupplierExclusion_0_4_Dow: x_Dow_4 = 0
SupplierExclusion_0_4_Huntsman: x_Huntsman_4 = 0
SupplierExclusion_0_4_Wanhua: x_Wanhua_4 = 0
SupplierExclusion_0_5_BASF: x_BASF_5 = 0
SupplierExclusion_0_5_Covestro: x_Covestro_5 = 0
SupplierExclusion_0_5_Huntsman: x_Huntsman_5 = 0
SupplierExclusion_0_5_Wanhua: x_Wanhua_5 = 0
TransitionLower_1_BASF: x_BASF_1 >= 0
TransitionLower_1_Covestro: x_Covestro_1 >= 0
TransitionLower_1_Huntsman: x_Huntsman_1 >= 0
TransitionLower_1_Wanhua: x_Wanhua_1 >= 0
TransitionLower_2_Covestro: x_Covestro_2 >= 0
TransitionLower_2_Dow: x_Dow_2 >= 0
TransitionLower_2_Huntsman: x_Huntsman_2 >= 0
TransitionLower_2_Wanhua: x_Wanhua_2 >= 0
TransitionLower_3_Covestro: x_Covestro_3 >= 0
TransitionLower_3_Dow: x_Dow_3 >= 0
TransitionLower_3_Huntsman: x_Huntsman_3 >= 0
TransitionLower_3_Wanhua: x_Wanhua_3 >= 0
TransitionLower_4_Covestro: x_Covestro_4 >= 0
TransitionLower_4_Dow: x_Dow_4 >= 0
TransitionLower_4_Huntsman: x_Huntsman_4 >= 0
TransitionLower_4_Wanhua: x_Wanhua_4 >= 0
TransitionLower_5_BASF: x_BASF_5 >= 0
TransitionLower_5_Covestro: x_Covestro_5 >= 0
TransitionLower_5_Huntsman: x_Huntsman_5 >= 0
TransitionLower_5_Wanhua: x_Wanhua_5 >= 0
Transition_1_BASF: - 8860493 T_1_BASF + x_BASF_1 <= 0
Transition_1_Covestro: - 8860493 T_1_Covestro + x_Covestro_1 <= 0
Transition_1_Huntsman: - 8860493 T_1_Huntsman + x_Huntsman_1 <= 0
Transition_1_Wanhua: - 8860493 T_1_Wanhua + x_Wanhua_1 <= 0
Transition_2_Covestro: - 7093293 T_2_Covestro + x_Covestro_2 <= 0
Transition_2_Dow: - 7093293 T_2_Dow + x_Dow_2 <= 0
Transition_2_Huntsman: - 7093293 T_2_Huntsman + x_Huntsman_2 <= 0
Transition_2_Wanhua: - 7093293 T_2_Wanhua + x_Wanhua_2 <= 0
Transition_3_Covestro: - 7423090 T_3_Covestro + x_Covestro_3 <= 0
Transition_3_Dow: - 7423090 T_3_Dow + x_Dow_3 <= 0
Transition_3_Huntsman: - 7423090 T_3_Huntsman + x_Huntsman_3 <= 0
Transition_3_Wanhua: - 7423090 T_3_Wanhua + x_Wanhua_3 <= 0
Transition_4_Covestro: - 14146424 T_4_Covestro + x_Covestro_4 <= 0
Transition_4_Dow: - 14146424 T_4_Dow + x_Dow_4 <= 0
Transition_4_Huntsman: - 14146424 T_4_Huntsman + x_Huntsman_4 <= 0
Transition_4_Wanhua: - 14146424 T_4_Wanhua + x_Wanhua_4 <= 0
Transition_5_BASF: - 4943948 T_5_BASF + x_BASF_5 <= 0
Transition_5_Covestro: - 4943948 T_5_Covestro + x_Covestro_5 <= 0
Transition_5_Huntsman: - 4943948 T_5_Huntsman + x_Huntsman_5 <= 0
Transition_5_Wanhua: - 4943948 T_5_Wanhua + x_Wanhua_5 <= 0
Volume_BASF: V_BASF - x_BASF_1 - x_BASF_2 - x_BASF_3 - x_BASF_4 - x_BASF_5 = 0
Volume_Covestro: V_Covestro - x_Covestro_1 - x_Covestro_2 - x_Covestro_3
 - x_Covestro_4 - x_Covestro_5 = 0
Volume_Dow: V_Dow - x_Dow_1 - x_Dow_2 - x_Dow_3 - x_Dow_4 - x_Dow_5 = 0
Volume_Huntsman: V_Huntsman - x_Huntsman_1 - x_Huntsman_2 - x_Huntsman_3
 - x_Huntsman_4 - x_Huntsman_5 = 0
Volume_Wanhua: V_Wanhua - x_Wanhua_1 - x_Wanhua_2 - x_Wanhua_3 - x_Wanhua_4
 - x_Wanhua_5 = 0
Bounds
 __dummy = 0
Binaries
T_1_BASF
T_1_Covestro
T_1_Huntsman
T_1_Wanhua
T_2_Covestro
T_2_Dow
T_2_Huntsman
T_2_Wanhua
T_3_Covestro
T_3_Dow
T_3_Huntsman
T_3_Wanhua
T_4_Covestro
T_4_Dow
T_4_Huntsman
T_4_Wanhua
T_5_BASF
T_5_Covestro
T_5_Huntsman
T_5_Wanhua
y_rebate_BASF_0
y_rebate_BASF_1
y_rebate_Covestro_0
y_rebate_Covestro_1
y_rebate_Huntsman_0
y_rebate_Huntsman_1
z_BASF
z_Covestro
z_Dow
z_Huntsman
z_Wanhua
End
