# -VB-AirWar
* Enemy aircraft manufacturing process(敌机制造流程)：

* 1 [AirWar_var_public]-[FPg_Add]-Case <New Plane ID>

* 2 [AirWar_var_public]-[Bg_Add]-Case <New Plane'Bullet ID(the same as New Plane ID)>

* 3 [AirWar_ruler_engine_Move]-[FoePlaneWYKZ]-Case <New Plane ID> 'It may be necessary to add new variables and new controls to achieve this function.(这里或许需要增加一些实现功能的新变量以及新控件)

* 4 [Form1]-Dim <New Plane's Name + "Time"> as Long

* 5 [Form1]-[Timer5_Timer] 'Setting new plane refresh time(这里设置新敌机的刷新间隔)

* 6 [AirWar_World_Engine]-[F5_Foe_Plane]-Case <New Plane ID>

* 7 [AirWar_ruler_engine_Phy]-[物理事件检测_Plane主体_PaB]-Case <New Plane ID>

* 8 [AirWar_World_Engine]-[F5_Foe_Bullet]-Case <New Plane'Bullet ID>

* 9 [AirWar_ruler_engine_Phy]-[物理事件检测_Plane主体_PBaFP事件_经验结算]-Case <New Plane ID>

* Ultimate skill manufacturing process（终极技能制造流程）：

* 1 [Form1]-Sk0n1(PCID*5+SKillID) Add new skill name controls

* 2 [Form1]-[Time2_Timer]-For c=1 to <SkillSum+1(SkillID)>

* 3 [AirWar_var_public]-Public PSkill(<PlayerSum>, <SkillSum+2(SkillID+1)>)

* 4 [AirWar_ruler_engine_Phy]-[物理事件检测_Plane主体_PBaFP事件_经验结算] 'Is it set here to unlock the enemy(这里设置是否为打败敌人解锁)

* 4.5 [Airwar_World_Case]-[升级] 'Here is set whether to unlock the upgrade(这里设置是否为升级解锁)

* 5 [Airwar_var_public]-[PBg_Add]-Case <SkillID>

* 6 [AirWar_var_public]-[PBg_Add_ClassificationAndEntry]-Case <SkillID> 'Set the specific properties of the skills here(这里设置技能具体属性)

* 7 [Airwar_ruler_engine_Move]-[PlBtWYKZ]-Case <SkillID> 'Set the skill run path here(这里设置技能运行路径)

* 8 [Airwar_ruler_engine]-[物理事件检测_PlaneBullet主体] 'Here is the entry of the skill physical function(这里设置技能物理函数入口)

* 9 [Airwar_ruler_engine_Phy]-Add New Functiion:物理事件检测_PlaneBullet主体_Trl_<SkillID+1> 'Set the physical effects in the function(函数内设置技能物理效果)

* 10 [Airwar_World_Engine]-[F5_PC_Bullet]-Case <SKillID> 'Setting the image effect of the skill(设置技能的图像效果)
