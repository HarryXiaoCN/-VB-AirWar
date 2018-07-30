Attribute VB_Name = "Airwar_ruler_Type"
Public Type Plane
    a As Boolean
    Da As Boolean
    HP As Long
    MxHP As Long
    Ar As Long '飞机物理半径
    E As Long '技能释放的能量
    MxE As Long
    Sb As Boolean '飞机是否会被击中
    Blt As Long '子弹类型
    DrSly As Long '死亡之后的补给类型
    Sp As Long '速度
    Emp As Long '经验值
    MxEmp As Long
    Rank As Long '等级
    Esp As Single  '能量恢复速度
    AiRank As Long 'Ai等级
    X As Single
    Y As Single
    mX As Single
    mY As Single
    dX As Single
    dY As Single
End Type
Public Type Bullet
    a As Boolean
    Da As Boolean
    Ar As Long
    Atk As Long '子弹的破坏力
    Pen As Boolean '是否有穿透性
    Trl As Long '子弹的轨迹类型
    Sb As Boolean '子弹是否会被破坏
    Source As Long '谁发出的
    Target As Long '目标
    Sp As Long
    X As Single
    Y As Single
    mX As Single
    mY As Single
    dX As Single
    dY As Single
End Type
Public Type Supply
    a As Boolean
    Da As Boolean
    Tp As Long '补给类型
    Sp As Long
    X As Single
    Y As Single
    mX As Single
    mY As Single
    dX As Single
    dY As Single
End Type
Public Type Skill
'    Public Type Skill_Bullet
        Blt As Long '子弹类型
        Exd As Long '能量消耗
'    Public Type Skill_Shield
        dHp As Long '生命加成
        dSb As Boolean '是否改变存在性
        CD As Long '冷却时间
'    Public Type Skill_Buff
        dAk As Long '攻击加成
        dE As Long '能量加成
End Type
