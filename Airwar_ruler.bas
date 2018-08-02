Attribute VB_Name = "Airwar_ruler_Type"
Public Type Plane
    a As Boolean
    Da As Boolean
    HP As Single
    MxHP As Single
    Ar As Long '飞机物理半径
    E As Long '技能释放的能量
    MxE As Long
    Sb As Boolean '飞机是否会被击中
    Blt As Long '子弹类型
    DrSly As Long '死亡之后的补给类型
    Sp As Single  '速度
    EMP As Long '经验值
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
    Atk As Single  '子弹的破坏力
    Pen As Boolean '是否有穿透性
    PenHp As Long
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
Public Type KeyConfig
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Attack As Integer
    Ultimate_Skill As Integer
    Skill_Switch As Integer
End Type
