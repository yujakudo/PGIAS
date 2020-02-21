Attribute VB_Name = "Constants"

'   �萔�B�V�[�g����񖼂Ȃ�
'   ���
Public Const C_IdGym As Integer = 0
Public Const C_IdMtc As Integer = 1
Public Const C_IdNormalAtk As Integer = 0
Public Const C_IdSpecialAtk As Integer = 1

Public Const C_TYPE As String = "�^�C�v"
Public Const C_SpeciesName As String = "�푰��"
Public Const C_NormalAttack As String = "�ʏ�킴"
Public Const C_SpecialAttack As String = "�Q�[�W�킴"
Public Const C_Self = "��"
Public Const C_Enemy = "�G"
Public Const C_Attack = "�U��"
Public Const C_Defense = "�h��"
Public Const C_Up = "��"
Public Const C_Down = "��"
Public Const C_LabelAlign = "���x������"
Public Const C_XAxis = "X��"
Public Const C_YAxis = "Y��"
Public Const C_XPrediction = "X���\��"
Public Const C_YPrediction = "Y���\��"

Public Const C_Prediction As String = "�\��"
Public Const C_Gym As String = "�W��"
Public Const C_Match As String = "�ΐ�"

Public Const TBL_NormalAtk As String = "�ʏ�킴�\"
Public Const TBL_SpecialAtk As String = "�Q�[�W�킴�\"

'   ��
'Public Const SH_Individual As String = "��"
Public Const IND_Nickname As String = "�j�b�N�l�[��"
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

Public Const IND_Predict As String = "_pr"
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
Public Const R_IndivMapTable As String = "�̃}�b�v���\"
Public Const R_IndivMapTypeSelect As String = "�̃}�b�v�^�C�v�I��"
Public Const R_IndivMapIndivSelect As String = "�̃}�b�v�̑I��"
Public Const R_IndivMapSettings As String = "�̃}�b�v�ݒ�"

Public Const IMAP_Name As String = "�j�b�N�l�[��"
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
Public Const ATK_Type As String = C_TYPE
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
Public Const ATK_GaugeNumber As String = "��_gage"
Public Const ATK_GaugeVolume As String = "��_gage"
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
Public Const CBT_Type As String = C_TYPE
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

'   �W���E�ΐ�
Public Const BE_R_Settngs As String = "�ݒ�"
Public Const ME_R_Settngs As String = "�ΐ�o�g���ݒ�"
Public Const BE_SetMode As String = "���[�h"
Public Const BE_SetSelfAtkDelay As String = "���U���x��"
Public Const BE_SetEnemyAtkDelay As String = "�G�U���x��"
Public Const BE_SetRankNum As String = "���ʐ�"
Public Const BE_SetRankVar As String = "���ʕt��"
Public Const BE_SetWithLimit As String = "����Z"
Public Const BE_DefCpUpper As String = "CP����f�t�H���g"
Public Const BE_DefCpLower As String = "CP�����f�t�H���g"

Public Const BE_R_DummyEnemy As String = "�_�~�["

Public Const BE_Species As String = C_SpeciesName
Public Const BE_Memo As String = "���l"
Public Const BE_NormalAttack As String = C_NormalAttack
Public Const BE_SpecialAttack As String = C_SpecialAttack
Public Const BE_SpecInput = "PL�̒l"
Public Const BE_PL = "PL"
Public Const BE_ATK = "ATK"
Public Const BE_DEF = "DEF"
Public Const BE_HP = "HP"
Public Const BE_CPHP = "CP/HP"
'Public Const BE_HP = "HP"
Public Const BE_UpperCP = "���"
Public Const BE_LowerCP = "����"
Public Const BE_BaseRank = "�V�󖢐ݒ�"
Public Const BE_SubRank = "�V��ݒ�"
Public Const BE_Weather As String = "�V��"
Public Const BE_Rank As String = "����"
Public Const BE_CtrName As String = "�j�b�N�l�[��"
Public Const BE_CtrAttack As String = C_SpecialAttack
Public Const BE_KT As String = "KT"
Public Const BE_KTR As String = "KTR"
Public Const BE_RankOut As String = "�����N����"
Public Const BE_CalcTime As String = "�v�Z����"

Public Const BE_RankBase As String = "����_b"

Public Const BE_SuffixBase As String = "_b"
Public Const BE_SuffixWeather As String = "_w"
Public Const BE_SuffixPredictBase As String = "_pb"
Public Const BE_SuffixPredictWeather As String = "_pw"

Public Const BE_NewEntryColorIndex As Integer = 30
Public Const BE_ReEntryColorIndex As Integer = 38

'   ���Ҍ�
Public Const NE_Name As String = "�j�b�N�l�[��"
Public Const NE_EntryNum As String = "�o�ꐔ"
Public Const NE_Type As String = "�^�C�v"
Public Const NE_PL As String = "PL"
Public Const NE_prPL As String = "PL_pr"
Public Const NE_Candies As String = "�A��"
Public Const NE_Sands As String = "���̍�"

Public Const NER_CountLower As String = "�v����������"
Public Const NELower_GymNow As String = "�W������"
Public Const NELower_GymPr As String = "�W���\��"
Public Const NELower_MtcNow As String = "�ΐ팻��"
Public Const NELower_MtcPr As String = "�ΐ�\��"

'   ���X�g
Public Const LI_R_Select As String = "���X�g�I��"
Public Const LI_Category As String = "�J�e�S��"
Public Const LI_Note As String = "���l"
Public Const LI_Species As String = C_SpeciesName
Public Const LI_PL As String = "PL"
Public Const LI_ATK As String = "ATK"
Public Const LI_DEF As String = "DEF"
Public Const LI_HP As String = "HP"
Public Const LI_CP As String = "CP"

Public Const RCT_L0 As String = "�T�J�L"
Public Const RCT_L1 As String = "�V�G��"
Public Const RCT_L2 As String = "�O���t"
Public Const RCT_L3 As String = "�A����"

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
Public Const R_RocketTroupe As String = "���P�b�g�c"
Public Const RT_Name As String = "����"
Public Const RT_Number As String = "�Ԗ�"
Public Const RT_Species As String = "�푰��"

'   �ő�l
Public Const R_DummyParameter As String = "�_�~�[�̓G�̃p�����[�^"
Public Const DM_AtkPower As String = "�U����"
Public Const DM_DefPower As String = "�h���"
Public Const DM_HP As String = "HP"
Public Const DM_GymNAtkPower As String = "�З�_ng"
Public Const DM_GymNAtkCharge As String = "�`���[�W_g"
Public Const DM_GymNAtkIdleTime As String = "��������"
Public Const DM_MtcNAtkPower As String = "�З�_nm"
Public Const DM_MtcNAtkCharge As String = "�`���[�W_m"
Public Const DM_MtcNAtkIdleTurn As String = "�����^�[��"
Public Const DM_GymSAtkPower As String = "�З�_s"
Public Const DM_GymSAtkGuageNum As String = "�Q�[�W��"
Public Const DM_GymSAtkIdleTime As String = "��������_s"
Public Const DM_MtcSAtkPower As String = "�З�_sm"
Public Const DM_MtcSAtkGuageVol As String = "�Q�[�W��"
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
Public Const msgConfirm As String = "�{����{0}���Ă������ł����H"
Public Const msgProcessing As String = "{1}��{0}���Ă��܂��B"
Public Const msgClear As String = "�N���A"
Public Const msgRemove As String = "�폜"
Public Const msgCalculate As String = "�v�Z"
Public Const msgRanking As String = "�����L���O"
Public Const msgExporting As String = "�G�N�X�|�[�g���Ă��܂��B"
Public Const msgImporting As String = "�C���|�[�g���Ă��܂��B"
Public Const msgAllSheet As String = "�S�ẴV�[�g"

Public Const msgMakingTable As String = "{0}�e�[�u�����쐬���Ă��܂��B"
Public Const msgMakingSheet As String = "{0}�V�[�g���쐬���Ă��܂��B"
Public Const msgMaking As String = "{0}���쐬���Ă��܂��B"
Public Const msgReseting As String = "{0}�����Z�b�g���Ă��܂��B"

Public Const msgSetColorToTypesOnTheSheet As String = "{0}�V�[�g�̃^�C�v�ɐF��t���Ă��܂��B"
Public Const msgSetColorToTypesAndAttcksOnTheSheet As String = "{0}�V�[�g�̃^�C�v�Ƃ킴���ɐF��t���Ă��܂��B"


Public Const msgAddAttackToSpecies As String = "�푰�V�[�g��{0}��{1}��{2}��ǉ����܂����B"
Public Const msgUnknownAttackName As String = "{1}��������܂���B{0}�V�[�g�ɒǉ����Ă��������B"
Public Const msgKeyDoesNotExist As String = "{0}�V�[�g��{1}��Ɂu{2}�v������܂���B"
Public Const msgColumnDoesNotExist As String = "{0}�V�[�g�Ɂu{1}�v�Ƃ�����͂���܂���B"
Public Const msgNoIdentifier As String = "{0}�V�[�g��{1}�s{2}��Ɂu{3}�v������܂���B"
Public Const msgColumnDoesNotExistOnTable As String = "�u{0}�v�Ɂu{1}�v�Ƃ�����͂���܂���B"

Public Const msgHitAttack As String = "{0}:{1}�B�_���[�W{2}�B"
Public Const msgMonStatus As String = "{0}: HP:{1}, Rsv:{2}"

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
'   Ranking
Public Const msgCalcRank As String = "{0}�̑΍􃉃��L���O���v�Z���Ă��܂��B"
Public Const msgSetWildCard As String = "���ׂĂ̑���Ƀ��C���h�J�[�h��ݒ肵�Ă��܂��B"
Public Const msgAddingRocketTroupe As String = "���P�b�g�c��ǉ����Ă��܂��B"
'   Controls
Public Const msgSureToAllClear As String = "���ׂăN���A���܂��B��낵���ł����H"
Public Const msgDoesOpenLog As String = "���O������܂��B�J���܂����H"
Public Const msgNoChange As String = "�ύX�͂���܂���B"
'   List
Public Const msgClearList As String = "���X�g���N���A���Ă��܂��B"




