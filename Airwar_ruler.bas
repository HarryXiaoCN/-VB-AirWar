Attribute VB_Name = "Airwar_ruler_Type"
Public Type Plane
    a As Boolean
    Da As Boolean
    HP As Single
    MxHP As Single
    Ar As Long '�ɻ�����뾶
    E As Long '�����ͷŵ�����
    MxE As Long
    Sb As Boolean '�ɻ��Ƿ�ᱻ����
    Blt As Long '�ӵ�����
    DrSly As Long '����֮��Ĳ�������
    Sp As Single  '�ٶ�
    EMP As Long '����ֵ
    MxEmp As Long
    Rank As Long '�ȼ�
    Esp As Single  '�����ָ��ٶ�
    AiRank As Long 'Ai�ȼ�
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
    Atk As Single  '�ӵ����ƻ���
    Pen As Boolean '�Ƿ��д�͸��
    PenHp As Long
    Trl As Long '�ӵ��Ĺ켣����
    Sb As Boolean '�ӵ��Ƿ�ᱻ�ƻ�
    Source As Long '˭������
    Target As Long 'Ŀ��
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
    Tp As Long '��������
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
