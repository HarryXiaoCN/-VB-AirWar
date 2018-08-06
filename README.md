# -VB-AirWar
Enemy aircraft manufacturing process(敌机制造流程)：

1 [AirWar_var_public]-[FPg_Add]-Case <New Plane ID>
  
2 [AirWar_var_public]-[Bg_Add]-Case <New Plane'Bullet ID(the same as New Plane ID)>

3 [AirWar_ruler_engine_Move]-[FoePlaneWYKZ]-Case <New Plane ID> 'It may be necessary to add new variables and new controls to achieve this function.(这里或许需要增加一些实现功能的新变量以及新控件)
  
4 [Form1]-Dim <New Plane's Name + "Time"> as Long

5 [Form1]-[Timer5] 'Setting new plane refresh time(这里设置新敌机的刷新间隔)

6 [AirWar_World_Engine]-[F5_Foe_Plane]-Case <New Plane ID>
  
7 [AirWar_ruler_engine_Phy]-[物理事件检测_Plane主体_PaB]-Case <New Plane ID>
  
8 [AirWar_World_Engine]-[F5_Foe_Bullet]-Case <New Plane'Bullet ID>

9 [AirWar_ruler_engine_Phy]-[物理事件检测_Plane主体_PBaFP事件_经验结算]-Case <New Plane ID>
