Attribute VB_Name = "Constants"

'   �萔�B�V�[�g����񖼂Ȃ�
'   ���
Public Const C_IdGym As Integer = 0
Public Const C_IdMtc As Integer = 1
Public Const C_IdNormalAtk As Integer = 0
Public Const C_IdSpecialAtk As Integer = 1
Public Const C_UpperCPl1 As Long = 1500
Public Const C_UpperCPl2 As Long = 2500
Public Const C_MaxPL As Long = 40
Public Const C_MaxLong As Long = 2000000000

Public Const C_Type As String = "�^�C�v"
Public Const C_SpeciesName As String = "�푰��"
Public Const C_Nickname As String = "�j�b�N�l�[��"
Public Const C_NormalAttack As String = "�ʏ�킴"
Public Const C_SpecialAttack As String = "�Q�[�W�킴"
Public Const C_Self = "��"
Public Const C_Enemy = "�G"
Public Const C_Attack = "�U��"
Public Const C_Defense = "�h��"
Public Const C_Up = "��"
Public Const C_Down = "��"
Public Const C_AutoTarget = "�����ڕW"
Public Const C_None = "�Ȃ�"
Public Const C_League1 = "�X�[�p�[���[�O"
Public Const C_League2 = "�n�C�p�[���[�O"
Public Const C_League3 = "�}�X�^�[���[�O"
Public Const C_Level = "���x��"
Public Const C_LabelAlign = "���x������"
Public Const C_XAxis = "X��"
Public Const C_YAxis = "Y��"
Public Const C_XPrediction = "X���\��"
Public Const C_YPrediction = "Y���\��"
Public Const C_CpUpper = "CP���"
Public Const C_PrCpLower = "�\��CP����"

Public Const C_Weather As String = "�V��"
Public Const C_Current As String = "����"
Public Const C_Prediction As String = "�\��"
Public Const C_Gym As String = "�W��"
Public Const C_Match As String = "�ΐ�"
Public Const C_SimMode As String = "Sim���[�h"
Public Const C_SelfAtkDelay As String = "���U���x��"
Public Const C_EnemyAtkDelay As String = "�G�U���x��"
Public Const C_KT As String = "KT"
Public Const C_KTR As String = "KTR"
Public Const C_ON As String = "ON"
Public Const C_OFF As String = "OFF"
Public Const C_Set As String = "�ݒ�"
Public Const C_NotSet As String = "���ݒ�"
Public Const C_DummyNormalAttack As String = "�_�~�[�ʏ�킴"
Public Const C_DummySpecialAttack As String = "�_�~�[�Q�[�W�킴"
Public Const C_Map As String = "�}�b�v"

Public Const cmdClear As String = "�N���A"
Public Const cmdRemove As String = "�폜"
Public Const cmdCalculate As String = "�v�Z"
Public Const cmdSetWeather As String = "�V��ݒ�"

Public Const TBL_NormalAtk As String = "�ʏ�킴�\"
Public Const TBL_SpecialAtk As String = "�Q�[�W�킴�\"

'   ��
'Public Const SH_Individual As String = "��"
Public Const IND_Nickname As String = C_Nickname
Public Const IND_Type1 As String = "�^�C�v1"
Public Const IND_Type2 As String = "�^�C�v2"
Public Const IND_Species As String = C_SpeciesName
Public Const IND_Number As String = "�ԍ�"
Public Const IND_CP As String = "CP"
Public Const IND_HP As String = "HP"
Public Const IND_indATK As String = "ATK_ind"
Public Const IND_indDEF As String = "DEF_ind"
Public Const IND_indHP As String = "HP_ind"
Public Const IND_NormalAtk As String = C_NormalAttack
Public Const IND_SpecialAtk1 As String = C_SpecialAttack & "1"
Public Const IND_SpecialAtk2 As String = C_SpecialAttack & "2"
Public Const IND_fixPL As String = "PL_fix"
Public Const IND_PL As String = "PL"
Public Const IND_League As String = "���[�O"
Public Const IND_AtkPower As String = "�U����"
Public Const IND_DefPower As String = "�h���"
Public Const IND_HP2 As String = "HP2"
Public Const IND_SCP As String = "SCP"
Public Const IND_DCP As String = "DCP"
Public Const IND_Endurance As String = "�ϋv��"

Public Const IND_GymBattle = "�@_g"
Public Const IND_MtcBattle = "�@_m"
Public Const IND_GymNormalAtkDamage As String = "�_���[�W_gn"
Public Const IND_GymNormalAtkTDPS As String = "tDPS_gn"
Public Const IND_GymSpecialAtk1Damage As String = "�_���[�W_gs1"
Public Const IND_GymSpecialAtk1TDPS As String = "tDPS_gs1"
Public Const IND_GymSpecialAtk1CDPS As String = "cDPS_gs1"
Public Const IND_GymSpecialAtk1Cycle As String = "Cyc_gs1"
Public Const IND_GymSpecialAtk2Damage As String = "�_���[�W_gs2"
Public Const IND_GymSpecialAtk2TDPS As String = "tDPS_gs2"
Public Const IND_GymSpecialAtk2Cycle As String = "Cyc_gs2"
Public Const IND_GymSpecialAtk2CDPS As String = "cDPS_gs2"
Public Const IND_MtcNormalAtkDamage As String = "�_���[�W_mn"
Public Const IND_MtcNormalAtkTDPS As String = "tDPT_mn"
Public Const IND_MtcSpecialAtk1Damage As String = "�_���[�W_ms1"
Public Const IND_MtcSpecialAtk1CDPS As String = "cDPT_ms1"
Public Const IND_MtcSpecialAtk1Cycle As String = "Cyc_ms1"
Public Const IND_MtcSpecialAtk2Damage As String = "�_���[�W_ms2"
Public Const IND_MtcSpecialAtk2CDPS As String = "cDPT_ms2"
Public Const IND_MtcSpecialAtk2Cycle As String = "Cyc_ms2"

Public Const IND_AutoTarget As String = "_pr"
Public Const IND_TargetPL As String = "PL_Target"
Public Const IND_prPL As String = "PL_pr"
Public Const IND_dPL As String = "��_PL"
Public Const IND_Candies As String = "�A��"
Public Const IND_Sands As String = "���̍�"
Public Const IND_prLeague As String = "���[�O_pr"
Public Const IND_prCP As String = "CP_pr"
Public Const IND_prHP As String = "HP_pr"
Public Const IND_DeltaHP As String = "��_HP"
Public Const IND_prAtkPower As String = "�U����_pr"
Public Const IND_DeltaAtkPower As String = "��_AtkP"
Public Const IND_prDefPower As String = "�h���_pr"
Public Const IND_DeltaDefPower As String = "��_DefP"
Public Const IND_prEndurance As String = "�ϋv��_pr"
Public Const IND_DeltaEndurance As String = "��_End"

Public Const IND_TargetNormalAtk As String = C_NormalAttack & "_pr"
Public Const IND_TargetSpecialAtk As String = C_SpecialAttack & "_pr"

Public Const IND_prGym As String = "�@_prg"
Public Const IND_prGymNormalAtkName As String = "�킴��_prgn"
Public Const IND_prGymNormalAtkDamage As String = "�_���[�W_prgn"
Public Const IND_DeltaGymNormalAtkDamage As String = "��_D_prgn"
Public Const IND_prGymNormalAtkTDPS As String = "tDPS_prgn"
Public Const IND_DeltaGymNormalAtkTDPS As String = "��_tDPS_prgn"
Public Const IND_prGymSpecialAtkName As String = "�킴��_prgs"
Public Const IND_prGymSpecialAtkDamage As String = "�_���[�W_prgs"
Public Const IND_DeltaGymSpecialAtkDamage As String = "��_D_prgs"
Public Const IND_prGymSpecialAtkTDPS As String = "tDPS_prgs"
Public Const IND_DeltaGymSpecialAtkTDPS As String = "��_tDPS_prgs"
Public Const IND_prGymCDpsNormalAtkName As String = "�ʏ�킴_cDPS_prg"
Public Const IND_prGymCDpsSpecialAtkName As String = "�Q�[�W�킴_cDPS_prg"
Public Const IND_prGymCDPS As String = "cDPS_prg"
Public Const IND_DeltaGymCDPS As String = "��_cDPS_prg"
Public Const IND_prGymCycle As String = "Cyc_prg"

Public Const IND_prMtc As String = "�@_prm"
Public Const IND_prMtcNormalAtkName As String = "�킴��_prmn"
Public Const IND_prMtcNormalAtkDamage As String = "�_���[�W_prmn"
Public Const IND_DeltaMtcNormalAtkDamage As String = "��_D_prmn"
Public Const IND_prMtcNormalAtkTDPS As String = "tDPT_prmn"
Public Const IND_DeltaMtcNormalAtkTDPS As String = "��_tDPT_prmn"
Public Const IND_prMtcSpecialAtkName As String = "�킴��_prms"
Public Const IND_prMtcSpecialAtkDamage As String = "�_���[�W_prms"
Public Const IND_DeltaMtcSpecialAtkDamage As String = "��_D_prms"
Public Const IND_prMtcCDpsNormalAtkName As String = "�ʏ�킴_cDPS_prm"
Public Const IND_prMtcCDpsSpecialAtkName As String = "�Q�[�W�킴_cDPS_prm"
Public Const IND_prMtcCDPS As String = "cDPT_prm"
Public Const IND_DeltaMtcCDPS As String = "��_cDPT_prm"
Public Const IND_prMtcCycle As String = "Cyc_prm"

'   �̃}�b�v
Public Const IMAP_R_Table As String = "�̃}�b�v���\"
Public Const IMAP_R_TypeSelect As String = "�̃}�b�v�^�C�v�I��"
Public Const IMAP_R_IndivSelect As String = "�̃}�b�v�̑I��"
Public Const IMAP_R_Settings As String = "�̃}�b�v�ݒ�"
Public Const IMAP_R_MakingTime As String = "�̃}�b�v�쐬����"

Public Const IMAP_Name As String = C_Nickname
Public Const IMAP_Species As String = C_SpeciesName
Public Const IMAP_Endurance As String = "�ϋv��"
Public Const IMAP_CDPS As String = "cDPS"
Public Const IMAP_isPrediction As String = "�\����"


'   �푰
Public Const R_SpeciesTable As String = "�푰���\"

Public Const SPEC_Number As String = "�ԍ�"
Public Const SPEC_Name As String = C_SpeciesName
Public Const SPEC_Type1 As String = "�^�C�v1"
Public Const SPEC_Type2 As String = "�^�C�v2"
Public Const SPEC_NormalAttack As String = C_NormalAttack
Public Const SPEC_NormalAttackLimited As String = "����" & C_NormalAttack
Public Const SPEC_SpecialAttack As String = C_SpecialAttack
Public Const SPEC_SpecialAttackLimited As String = "����" & C_SpecialAttack


'   �킴
Public Const ATK_Name As String = "�킴��"
Public Const ATK_Type As String = C_Type
Public Const ATK_GymPower As String = "�З�_g"
Public Const ATK_MtcPower As String = "�З�_m"
Public Const ATK_GymCharge As String = "�`���[�W_g"
Public Const ATK_MtcCharge As String = "�`���[�W_m"
Public Const ATK_IdleTime As String = "����_it"
Public Const ATK_DPS As String = "DPS"
Public Const ATK_EPS As String = "EPS"
Public Const ATK_DamageDelay As String = "����_dd"
Public Const ATK_IdleTurnNum As String = "�^�[��_it"
Public Const ATK_DPT As String = "DPT"
Public Const ATK_EPT As String = "EPT"
Public Const ATK_GaugeNumber As String = "��_gg"
Public Const ATK_GaugeVolume As String = "��_gg"
Public Const ATK_DPE As String = "DPE"
Public Const ATK_Effect As String = "����"
Public Const ATK_EffectStep As String = "�K"
Public Const ATK_EffectProb As String = "��"

Public Const ATK_typeMatch As String = "�␳"
Public Const ATK_CorrGymPower As String = "�З�_gc"
Public Const ATK_CorrDPS As String = "DPS_gc"
Public Const ATK_CorrMtcPower As String = "�З�_mc"
Public Const ATK_CorrDPT As String = "DPT_mc"
Public Const ATK_CorrDPE As String = "DPE_mc"

'   �^�C�v�ʃV�[�g
Public Const R_ClassifiedByType As String = "�^�C�v�ʕ\"
Public Const CBT_Type As String = C_Type
Public Const CBT_DoubleWeak As String = "x2.56"
Public Const CBT_SingleWeak As String = "x1.6"
Public Const CBT_Soso As String = "x1.0"
Public Const CBT_SingleToler As String = "x0.625"
Public Const CBT_DoubleToler As String = "x0.39"
Public Const CBT_OverToler As String = "x0.244"
Public Const CBT_NumOfSpecies As String = "�푰��"
Public Const CBT_Species As String = C_SpeciesName

'   �푰����1
Public Const R_SpeciesAnalysis1 As String = "�푰����1"
Public Const SA1_Number As String = "�ԍ�"
Public Const SA1_Name As String = C_SpeciesName
Public Const SA1_Type As String = "�^�C�v"
Public Const SA1_Weakness As String = "��_"
Public Const SA1_DoubleWeak As String = "x2.56"
Public Const SA1_SingleWeak As String = "x1.6"
Public Const SA1_Tolerance As String = "�ϐ�"
Public Const SA1_SingleToler As String = "x0.625"
Public Const SA1_DoubleToler As String = "x0.39"
Public Const SA1_OverToler As String = "x0.244"
Public Const SA1_Param As String = "�]���p�����[�^"
Public Const SA1_ATKPower As String = "�U����"
Public Const SA1_DEFPower As String = "�h���"
Public Const SA1_HP As String = "HP"
Public Const SA1_CP As String = "CP"
Public Const SA1_SCP As String = "SCP"
Public Const SA1_DCP As String = "DCP"
Public Const SA1_Endurance As String = "�ϋv��"

Public Const SA1_GymBattleT = "�W��"
Public Const SA1_MtcBattleT = "�ΐ�"
Public Const SA1_GymBattle = "�@_g"
Public Const SA1_MtcBattle = "�@_m"

Public Const SA1_NA_Damage = "�_���[�W�ő�" & C_NormalAttack
Public Const SA1_NA_DamageAtkName = "�킴��_gnD"
Public Const SA1_NA_DamageValue = "�_���[�W_gnD"
Public Const SA1_NA_Dps = "tDPS�ő�" & C_NormalAttack
Public Const SA1_NA_DpsAtkName = "�킴��_gnDps"
Public Const SA1_NA_DpsValue = "tDPS_gnDps"
Public Const SA1_NA_Eps = "�`���[�W�ő�" & C_NormalAttack
Public Const SA1_NA_EpsAtkName = "�킴��_gnEps"
Public Const SA1_NA_EpsValue = "EPS_gnEps"

Public Const SA1_SA_Damage = "�_���[�W�ő�" & C_SpecialAttack
Public Const SA1_SA_DamageAtkName = "�킴��_gsD"
Public Const SA1_SA_DamageValue = "�_���[�W_gsD"
Public Const SA1_SA_Dps = "tDPS�ő�" & C_SpecialAttack
Public Const SA1_SA_DpsAtkName = "�킴��_gsDps"
Public Const SA1_SA_DpsValue = "tDPS_gsDps"

Public Const SA1_cDpsRank = "cDPS�����N"
Public Const SA1_CDSP_NormalAtkName = "�ʏ�킴_gnCdps"
Public Const SA1_CDSP_SpecialAtkName = "�Q�[�W�킴_gsCdps"
Public Const SA1_CDSP_Value = "cDPS_gCdps"
Public Const SA1_CDSP_Cycle = "Cyc_gCdps"


Public Const SA1_CDST_NormalAtkName = "�ʏ�킴_mnCdps"
Public Const SA1_CDST_SpecialAtkName = "�Q�[�W�킴_msCdps"
Public Const SA1_CDST_Value = "cDPT_mCdps"

Public Const SA1_SA_Dpe = "tDPE�ő�" & C_SpecialAttack
Public Const SA1_SA_DpeAtkName = "�Q�[�WtDPE�Z"
Public Const SA1_SA_DpeValue = "�Q�[�WtDPE"

Public Const SA1_LastColumn As String = "��cDPT"

Public Const R_NormalAtkSpeciesSelect As String = "�ʏ�킴�푰�I��"
Public Const R_SpecialAtkSpeciesSelect As String = "�Q�[�W�킴�푰�I��"

'   �푰�}�b�v
Public Const R_SpeciesMapTypeSelect As String = "�푰�}�b�v�^�C�v�I��"
Public Const R_SpeciesMapSpeciesSelect As String = "�푰�}�b�v�푰�I��"
Public Const R_SpeciesMapSettings As String = "�푰�}�b�v�ݒ�"

'   �΍􃉃��N
Public Const CR_Weather As String = C_Weather
Public Const CR_Species As String = C_SpeciesName
Public Const CR_Memo As String = "���l"
Public Const CR_Attacks As String = "/�Q�[�W�킴"
Public Const CR_PL As String = "PL"
Public Const CR_ATK As String = "ATK"
Public Const CR_DEF As String = "DEF"
Public Const CR_HP As String = "HP"
Public Const CR_CPHP As String = "CP/HP"
Public Const CR_CPLimit As String = "/����"
Public Const CR_Time As String = "/����"

Public Const CR_Current As String = "����"
Public Const CR_Prediction As String = "�\��"

Public Const CR_Rank As String = "����"
Public Const CR_CtrName As String = C_Nickname
Public Const CR_CtrPL As String = "PL"
Public Const CR_CtrNormalAttack As String = C_NormalAttack
Public Const CR_CtrSpecialAttack As String = C_SpecialAttack
Public Const CR_CtrCDPS As String = "cDPS"

Public Const CR_KT As String = C_KT
Public Const CR_KTR As String = C_KTR

Public Const CR_SuffixBase As String = "_b"
Public Const CR_SuffixPredict As String = "_p"
Public Const CR_SuffixWeather As String = "_w"

Public Const CR_NewEntryColorIndex As Integer = 30
Public Const CR_ReEntryColorIndex As Integer = 38
Public Const CR_DropEntryColorIndex As Integer = 23
'   �ݒ�
Public Const CR_R_ListSelect As String = "�΍􃊃X�g�I��"
Public Const CR_R_AllCalcTime As String = "�΍��S�v�Z����"
Public Const CR_R_WeatherGuess As String = "�΍��V��ݒ�"
Public Const CR_R_Settngs As String = "�΍��ݒ�"
Public Const CR_SetMode As String = C_SimMode
Public Const CR_SetSelfAtkDelay As String = C_SelfAtkDelay
Public Const CR_SetEnemyAtkDelay As String = C_EnemyAtkDelay
Public Const CR_SetRankNum As String = "���ʐ�"
Public Const CR_SetRankVar As String = "���ʕt��"
Public Const CR_SetWithLimit_b As String = "�󔒎�����Z_b"
Public Const CR_DefCpUpper As String = "CP����f�t�H���g"
Public Const CR_DefCpLower As String = "CP�����f�t�H���g"
Public Const CR_CountRankCur As String = "���ݏW�v�����N"
Public Const CR_CountRankPr As String = "�\���W�v�����N"

'   �_�~�[�̐ݒ�
Public Const CR_R_DummyEnemy As String = "�΍�_�~�[�ݒ�"
Public Const CR_DmyAtkPower As String = "�U����"
Public Const CR_DmyDefPower As String = "�h���"
Public Const CR_DmyHP As String = "HP"
Public Const CR_DmyCP As String = "CP"

Public Const CR_SheetPrefix As String = "�΍�"

'   �W�v
Public Const NE_Name As String = C_Nickname
Public Const NE_EntryNum As String = "�o�ꐔ"
Public Const NE_FlagedNum As String = "�o�ꐔ_f"
Public Const NE_Type As String = "�^�C�v"
Public Const NE_PL As String = "PL"
Public Const NE_prPL As String = "PL_pr"
Public Const NE_Candies As String = "�A��"
Public Const NE_Sands As String = "���̍�"
Public Const NE_ColumnsNum As Integer = 8
Public Const NE_DataRow As Long = 5

Public Const NE_TableName As String = "�W�v�\"
Public Const NE_CalcAllTime As String = "�W�v�S�v�Z����"

'   ���X�g
Public Const LI_R_Select As String = "���X�g�I��"
Public Const LI_R_Command As String = "���X�g�R�}���h"
Public Const LI_Category As String = "�J�e�S��"
Public Const LI_Note As String = "���l"
Public Const LI_Species As String = C_SpeciesName
Public Const LI_PL As String = "PL"
Public Const LI_ATK As String = "ATK"
Public Const LI_DEF As String = "DEF"
Public Const LI_HP As String = "HP"
Public Const LI_CP As String = "CP"
Public Const LI_DefaultListName As String = "��`���X�g"
'   �R�}���h
Public Const LI_CMD_Clear As String = "�N���A"
Public Const LI_CMD_SetAsRocket As String = "���P�b�g�c�p�����[�^�ݒ�"
'   ���[�_�[��
Public Const RCT_L0 As String = "�T�J�L"
Public Const RCT_L1 As String = "�V�G��"
Public Const RCT_L2 As String = "�O���t"
Public Const RCT_L3 As String = "�A����"

'   �Q��
Public Const RFR_Nickname As String = C_Nickname
Public Const RFR_Type As String = C_Type
Public Const RFR_NormalAtk As String = C_NormalAttack
Public Const RFR_SpecialAtk As String = C_SpecialAttack


'   ����
Public Const R_Type As String = "�^�C�v"
Public Const R_InterTypeInflu As String = "�����\"
Public Const R_TypeSynonym As String = "�^�C�v����"
Public Const R_WeatherBoost As String = "�V��u�[�X�g"
Public Const R_interTypeFactor As String = "�^�C�v���֌W��"
Public Const R_WeatherFactor As String = "�V��u�[�X�g�W��"
Public Const R_TypeMatchFactor As String = "�^�C�v��v�W��"
Public Const R_MtcBtlFactor As String = "�g���[�i�[�o�g���W��"
Public Const R_ChargeByDamage As String = "��_���[�W�ɂ��`���[�W"
Public Const R_WhiteSpace As String = "��"
Public Const R_WeatherTable As String = "�V��\"

'   CPM
Public Const R_StatusTransition As String = "�킴�̔\�͕ω�"
'Public Const R_RocketTroupe As String = "���P�b�g�c"
'Public Const RT_Name As String = "����"
'Public Const RT_Number As String = "�Ԗ�"
'Public Const RT_Species As String = "�푰��"

'   �ő�l
Public Const R_DummyParameter As String = "�_�~�[�̓G�̃p�����[�^"
'   Ver
Public Const VH_Branch As String = "�u�����`"
Public Const VH_Version As String = "Ver."
Public Const VH_Date As String = "���t"
Public Const VH_Article As String = "�L��"

Public Const R_GlobalSettings As String = "�S�̐ݒ�"
Public Const GS_FileName As String = "�t�@�C����"
Public Const GS_BranchName As String = "�u�����`��"
Public Const GS_DistDir As String = "�f�B���N�g��"
Public Const GS_UseLimitedAttacksOnSpeciesAna As String = "�푰���͂Ō���킴"


'   ���b�Z�[�W
Public Const msgDoSomething As String = "..."
Public Const msgConfirm As String = "�{����{0}���Ă������ł����H"
Public Const msgProcessing As String = "{1}��{0}���Ă��܂��B"
Public Const msgRanking As String = "�����L���O"
Public Const msgExporting As String = "�G�N�X�|�[�g���Ă��܂��B"
Public Const msgImporting As String = "�C���|�[�g���Ă��܂��B"
Public Const msgAllSheet As String = "�S�ẴV�[�g"
Public Const msgAllInDivAna As String = "�S�Ă̌̕���"
Public Const msgChaingingSettings As String = "�ݒ��ύX���Ă��܂��B"

Public Const msgMakingTable As String = "{0}�e�[�u�����쐬���Ă��܂��B"
Public Const msgMakingSheet As String = "{0}�V�[�g���쐬���Ă��܂��B"
Public Const msgMaking As String = "{0}���쐬���Ă��܂��B"
Public Const msgReseting As String = "{0}�����Z�b�g���Ă��܂��B"
Public Const msgInitializing As String = "{0}�����������Ă��܂��B"

Public Const msgSetColorToTypesOnTheSheet As String = "{0}�V�[�g�̃^�C�v�ɐF��t���Ă��܂��B"
Public Const msgSetColorToTypesAndAttcksOnTheSheet As String = "{0}�V�[�g�̃^�C�v�Ƃ킴���ɐF��t���Ă��܂��B"


Public Const msgAddAttackToSpecies As String = "�푰�V�[�g��{0}��{1}��{2}��ǉ����܂����B"
Public Const msgUnknownAttackName As String = "{1}��������܂���B{0}�V�[�g�ɒǉ����Ă��������B"
Public Const msgKeyDoesNotExist As String = "{0}�V�[�g��{1}��Ɂu{2}�v������܂���B"
Public Const msgColumnDoesNotExist As String = "{0}�V�[�g�Ɂu{1}�v�Ƃ�����͂���܂���B"
Public Const msgNoIdentifier As String = "{0}�V�[�g��{1}�s{2}��Ɂu{3}�v������܂���B"
Public Const msgColumnDoesNotExistOnTable As String = "�u{0}�v�Ɂu{1}�v�Ƃ�����͂���܂���B"

'   shSpecies
Public Const msgAttackIsLimited As String = "�u{0}�v���푰�V�[�g�ɒǉ����܂��B�u{0}�v�͌���킴�ł����H" _
                & vbCrLf & "[�͂�]=����킴�A[������]=���ʂ̂킴�A[�L�����Z��]=�ǉ����Ȃ�"
'   shNormalAttack, shSpecialAttack
Public Const msgResetSelectedAttack As String = "�킴���̑I�������Z�b�g���Ă��܂��B"
'   shSpeciesAnalysis1
Public Const msgSetNewSpeciesToSpeciesAnalysis1Sheet As String = "�푰���̓V�[�g�Ɏ푰��ǉ����Ă��܂��"
'   shIndividual
Public Const msgCalculatingIndividualSheet As String = "�̃V�[�g�̉�͒l���v�Z���Ă��܂��B"
Public Const msgAligningIndividualSheet As String = "�̃V�[�g�𒲐����Ă��܂��B"
Public Const msgPLis0 As String = "PL��0�̌̂�����܂����B�p�����[�^���m�F���Ă��������B"
'   shIndivMap
Public Const msgAllRecalc As String = "�S�Ă̌̕��̓V�[�g�̍Čv�Z�����Ă��܂��B"
Public Const msgSettingMap As String = "�}�b�v��ݒ肵�Ă��܂��B"
'   Ranking
Public Const msgCopyingRegion As String = "�̈���R�s�[���Ă��܂��B"
Public Const msgCalcRank As String = "{0}�̑΍􃉃��L���O���v�Z���Ă��܂��B"
Public Const msgSetWildCard As String = "�_�~�[�̑����ݒ肵�Ă��܂��B"
Public Const msgAddingListItems As String = "���X�g�u{0}�v���瑊���ǉ����Ă��܂��B"
Public Const msgAskChangeBattleMode As String = "�ΐ탂�[�h��ύX����ƁA���ׂĂ̑�����폜���܂��B��낵���ł����H"
'   Controls
Public Const msgSureToAllClear As String = "���ׂăN���A���܂��B��낵���ł����H"
Public Const msgDoesOpenLog As String = "���O������܂��B�J���܂����H"
Public Const msgNoChange As String = "�ύX�͂���܂���B"
'   List
Public Const msgClearList As String = "���X�g {0} ���N���A���Ă��܂��B"

