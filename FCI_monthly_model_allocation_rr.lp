\* Output\FCI_monthly_model_allocation_rr *\
Minimize
OBJ: 1095.5 x_rice_('ARA',_'AKT') + 1897.2 x_rice_('ARA',_'BH')
 + 1209.9 x_rice_('HTW',_'AKT') + 1897.2 x_rice_('HTW',_'BH')
 + 571 x_wheat_('JHL',_'BPZ') + 2263.7 x_wheat_('JHL',_'FSNG')
 + 454.2 x_wheat_('PCWD',_'BPZ') + 2241.1 x_wheat_('PCWD',_'FSNG')
Subject To
_C1: x_wheat_('JHL',_'BPZ') + x_wheat_('JHL',_'FSNG') <= 1
_C2: x_wheat_('PCWD',_'BPZ') + x_wheat_('PCWD',_'FSNG') <= 1
_C3: x_rice_('ARA',_'AKT') + x_rice_('ARA',_'BH') <= 1
_C4: x_rice_('HTW',_'AKT') + x_rice_('HTW',_'BH') <= 1
_C5: x_wheat_('JHL',_'FSNG') + x_wheat_('PCWD',_'FSNG') >= 1
_C6: x_wheat_('JHL',_'BPZ') + x_wheat_('PCWD',_'BPZ') >= 1
_C7: x_rice_('ARA',_'AKT') + x_rice_('HTW',_'AKT') >= 1
_C8: x_rice_('ARA',_'BH') + x_rice_('HTW',_'BH') >= 1
End
