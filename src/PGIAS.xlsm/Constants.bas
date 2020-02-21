Attribute VB_Name = "Constants"

'   定数。シート名や列名など
'   一般
Public Const C_IdGym As Integer = 0
Public Const C_IdMtc As Integer = 1
Public Const C_IdNormalAtk As Integer = 0
Public Const C_IdSpecialAtk As Integer = 1

Public Const C_TYPE As String = "タイプ"
Public Const C_SpeciesName As String = "種族名"
Public Const C_NormalAttack As String = "通常わざ"
Public Const C_SpecialAttack As String = "ゲージわざ"
Public Const C_Self = "自"
Public Const C_Enemy = "敵"
Public Const C_Attack = "攻撃"
Public Const C_Defense = "防御"
Public Const C_Up = "↑"
Public Const C_Down = "↓"
Public Const C_LabelAlign = "ラベル調整"
Public Const C_XAxis = "X軸"
Public Const C_YAxis = "Y軸"
Public Const C_XPrediction = "X軸予測"
Public Const C_YPrediction = "Y軸予測"

Public Const C_Prediction As String = "予測"
Public Const C_Gym As String = "ジム"
Public Const C_Match As String = "対戦"

Public Const TBL_NormalAtk As String = "通常わざ表"
Public Const TBL_SpecialAtk As String = "ゲージわざ表"

'   個体
'Public Const SH_Individual As String = "個体"
Public Const IND_Nickname As String = "ニックネーム"
Public Const IND_Type1 As String = "タイプ1"
Public Const IND_Type2 As String = "タイプ2"
Public Const IND_Species As String = C_SpeciesName
Public Const IND_Number As String = "番号"
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
Public Const IND_League As String = "リーグ"
Public Const IND_AtkPower As String = "攻撃力"
Public Const IND_DefPower As String = "防御力"
Public Const IND_HP2 As String = "HP2"
Public Const IND_SCP As String = "SCP"
Public Const IND_DCP As String = "DCP"
Public Const IND_Endurance As String = "耐久力"

Public Const IND_GymBattle = "　_g"
Public Const IND_MtcBattle = "　_m"
Public Const IND_GymNormalAtkDamage As String = "ダメージ_gn"
Public Const IND_GymNormalAtkTDPS As String = "tDPS_gn"
Public Const IND_GymSpecialAtk1Damage As String = "ダメージ_gs1"
Public Const IND_GymSpecialAtk1TDPS As String = "tDPS_gs1"
Public Const IND_GymSpecialAtk1CDPS As String = "cDPS_gs1"
Public Const IND_GymSpecialAtk1Cycle As String = "Cyc_gs1"
Public Const IND_GymSpecialAtk2Damage As String = "ダメージ_gs2"
Public Const IND_GymSpecialAtk2TDPS As String = "tDPS_gs2"
Public Const IND_GymSpecialAtk2Cycle As String = "Cyc_gs2"
Public Const IND_GymSpecialAtk2CDPS As String = "cDPS_gs2"
Public Const IND_MtcNormalAtkDamage As String = "ダメージ_mn"
Public Const IND_MtcNormalAtkTDPS As String = "tDPT_mn"
Public Const IND_MtcSpecialAtk1Damage As String = "ダメージ_ms1"
Public Const IND_MtcSpecialAtk1CDPS As String = "cDPT_ms1"
Public Const IND_MtcSpecialAtk1Cycle As String = "Cyc_ms1"
Public Const IND_MtcSpecialAtk2Damage As String = "ダメージ_ms2"
Public Const IND_MtcSpecialAtk2CDPS As String = "cDPT_ms2"
Public Const IND_MtcSpecialAtk2Cycle As String = "Cyc_ms2"

Public Const IND_Predict As String = "_pr"
Public Const IND_TargetPL As String = "PL_Target"
Public Const IND_prPL As String = "PL_pr"
Public Const IND_dPL As String = "Δ_PL"
Public Const IND_Candies As String = "アメ"
Public Const IND_Sands As String = "星の砂"
Public Const IND_prLeague As String = "リーグ_pr"
Public Const IND_prCP As String = "CP_pr"
Public Const IND_prHP As String = "HP_pr"
Public Const IND_DeltaHP As String = "Δ_HP"
Public Const IND_prAtkPower As String = "攻撃力_pr"
Public Const IND_DeltaAtkPower As String = "Δ_AtkP"
Public Const IND_prDefPower As String = "防御力_pr"
Public Const IND_DeltaDefPower As String = "Δ_DefP"
Public Const IND_prEndurance As String = "耐久力_pr"
Public Const IND_DeltaEndurance As String = "Δ_End"

Public Const IND_TargetNormalAtk As String = C_NormalAttack & "_pr"
Public Const IND_TargetSpecialAtk As String = C_SpecialAttack & "_pr"

Public Const IND_prGym As String = "　_prg"
Public Const IND_prGymNormalAtkName As String = "わざ名_prgn"
Public Const IND_prGymNormalAtkDamage As String = "ダメージ_prgn"
Public Const IND_DeltaGymNormalAtkDamage As String = "Δ_D_prgn"
Public Const IND_prGymNormalAtkTDPS As String = "tDPS_prgn"
Public Const IND_DeltaGymNormalAtkTDPS As String = "Δ_tDPS_prgn"
Public Const IND_prGymSpecialAtkName As String = "わざ名_prgs"
Public Const IND_prGymSpecialAtkDamage As String = "ダメージ_prgs"
Public Const IND_DeltaGymSpecialAtkDamage As String = "Δ_D_prgs"
Public Const IND_prGymSpecialAtkTDPS As String = "tDPS_prgs"
Public Const IND_DeltaGymSpecialAtkTDPS As String = "Δ_tDPS_prgs"
Public Const IND_prGymCDpsNormalAtkName As String = "通常わざ_cDPS_prg"
Public Const IND_prGymCDpsSpecialAtkName As String = "ゲージわざ_cDPS_prg"
Public Const IND_prGymCDPS As String = "cDPS_prg"
Public Const IND_DeltaGymCDPS As String = "Δ_cDPS_prg"
Public Const IND_prGymCycle As String = "Cyc_prg"

Public Const IND_prMtc As String = "　_prm"
Public Const IND_prMtcNormalAtkName As String = "わざ名_prmn"
Public Const IND_prMtcNormalAtkDamage As String = "ダメージ_prmn"
Public Const IND_DeltaMtcNormalAtkDamage As String = "Δ_D_prmn"
Public Const IND_prMtcNormalAtkTDPS As String = "tDPT_prmn"
Public Const IND_DeltaMtcNormalAtkTDPS As String = "Δ_tDPT_prmn"
Public Const IND_prMtcSpecialAtkName As String = "わざ名_prms"
Public Const IND_prMtcSpecialAtkDamage As String = "ダメージ_prms"
Public Const IND_DeltaMtcSpecialAtkDamage As String = "Δ_D_prms"
Public Const IND_prMtcCDpsNormalAtkName As String = "通常わざ_cDPS_prm"
Public Const IND_prMtcCDpsSpecialAtkName As String = "ゲージわざ_cDPS_prm"
Public Const IND_prMtcCDPS As String = "cDPT_prm"
Public Const IND_DeltaMtcCDPS As String = "Δ_cDPT_prm"
Public Const IND_prMtcCycle As String = "Cyc_prm"

'   個体マップ
Public Const R_IndivMapTable As String = "個体マップ元表"
Public Const R_IndivMapTypeSelect As String = "個体マップタイプ選択"
Public Const R_IndivMapIndivSelect As String = "個体マップ個体選択"
Public Const R_IndivMapSettings As String = "個体マップ設定"

Public Const IMAP_Name As String = "ニックネーム"
Public Const IMAP_Species As String = C_SpeciesName
Public Const IMAP_Endurance As String = "耐久力"
Public Const IMAP_CDPS As String = "cDPS"
Public Const IMAP_isPrediction As String = "予測か"


'   種族
Public Const R_SpeciesTable As String = "種族元表"

Public Const SPEC_Number As String = "番号"
Public Const SPEC_Name As String = C_SpeciesName
Public Const SPEC_Type1 As String = "タイプ1"
Public Const SPEC_Type2 As String = "タイプ2"
Public Const SPEC_NormalAttack As String = C_NormalAttack
Public Const SPEC_NormalAttackLimited As String = "限定" & C_NormalAttack
Public Const SPEC_SpecialAttack As String = C_SpecialAttack
Public Const SPEC_SpecialAttackLimited As String = "限定" & C_SpecialAttack


'   わざ
Public Const ATK_Name As String = "わざ名"
Public Const ATK_Type As String = C_TYPE
Public Const ATK_GymPower As String = "威力_g"
Public Const ATK_MtcPower As String = "威力_m"
Public Const ATK_GymCharge As String = "チャージ_g"
Public Const ATK_MtcCharge As String = "チャージ_m"
Public Const ATK_IdleTime As String = "時間_it"
Public Const ATK_DPS As String = "DPS"
Public Const ATK_EPS As String = "EPS"
Public Const ATK_DamageDelay As String = "時間_dd"
Public Const ATK_IdleTurnNum As String = "ターン_it"
Public Const ATK_DPT As String = "DPT"
Public Const ATK_EPT As String = "EPT"
Public Const ATK_GaugeNumber As String = "数_gage"
Public Const ATK_GaugeVolume As String = "量_gage"
Public Const ATK_DPE As String = "DPE"
Public Const ATK_Effect As String = "効果"
Public Const ATK_EffectStep As String = "階"
Public Const ATK_EffectProb As String = "率"

Public Const ATK_typeMatch As String = "補正"
Public Const ATK_CorrGymPower As String = "威力_gc"
Public Const ATK_CorrDPS As String = "DPS_gc"
Public Const ATK_CorrMtcPower As String = "威力_mc"
Public Const ATK_CorrDPT As String = "DPT_mc"
Public Const ATK_CorrDPE As String = "DPE_mc"

'   タイプ別シート
Public Const R_ClassifiedByType As String = "タイプ別表"
Public Const CBT_Type As String = C_TYPE
Public Const CBT_DoubleWeak As String = "x2.56"
Public Const CBT_SingleWeak As String = "x1.6"
Public Const CBT_Soso As String = "x1.0"
Public Const CBT_SingleToler As String = "x0.625"
Public Const CBT_DoubleToler As String = "x0.39"
Public Const CBT_OverToler As String = "x0.244"
Public Const CBT_NumOfSpecies As String = "種族数"
Public Const CBT_Species As String = C_SpeciesName

'   種族分析1
Public Const R_SpeciesAnalysis1 As String = "種族分析1"
Public Const SA1_Number As String = "番号"
Public Const SA1_Name As String = C_SpeciesName
Public Const SA1_Type As String = "タイプ"
Public Const SA1_Weakness As String = "弱点"
Public Const SA1_DoubleWeak As String = "x2.56"
Public Const SA1_SingleWeak As String = "x1.6"
Public Const SA1_Tolerance As String = "耐性"
Public Const SA1_SingleToler As String = "x0.625"
Public Const SA1_DoubleToler As String = "x0.39"
Public Const SA1_OverToler As String = "x0.244"
Public Const SA1_Param As String = "評価パラメータ"
Public Const SA1_ATKPower As String = "攻撃力"
Public Const SA1_DEFPower As String = "防御力"
Public Const SA1_HP As String = "HP"
Public Const SA1_CP As String = "CP"
Public Const SA1_SCP As String = "SCP"
Public Const SA1_DCP As String = "DCP"
Public Const SA1_Endurance As String = "耐久力"

Public Const SA1_GymBattleT = "ジム"
Public Const SA1_MtcBattleT = "対戦"
Public Const SA1_GymBattle = "　_g"
Public Const SA1_MtcBattle = "　_m"

Public Const SA1_NA_Damage = "ダメージ最大" & C_NormalAttack
Public Const SA1_NA_DamageAtkName = "わざ名_gnD"
Public Const SA1_NA_DamageValue = "ダメージ_gnD"
Public Const SA1_NA_Dps = "tDPS最大" & C_NormalAttack
Public Const SA1_NA_DpsAtkName = "わざ名_gnDps"
Public Const SA1_NA_DpsValue = "tDPS_gnDps"
Public Const SA1_NA_Eps = "チャージ最速" & C_NormalAttack
Public Const SA1_NA_EpsAtkName = "わざ名_gnEps"
Public Const SA1_NA_EpsValue = "EPS_gnEps"

Public Const SA1_SA_Damage = "ダメージ最大" & C_SpecialAttack
Public Const SA1_SA_DamageAtkName = "わざ名_gsD"
Public Const SA1_SA_DamageValue = "ダメージ_gsD"
Public Const SA1_SA_Dps = "tDPS最大" & C_SpecialAttack
Public Const SA1_SA_DpsAtkName = "わざ名_gsDps"
Public Const SA1_SA_DpsValue = "tDPS_gsDps"

Public Const SA1_cDpsRank = "cDPSランク"
Public Const SA1_CDSP_NormalAtkName = "通常わざ_gnCdps"
Public Const SA1_CDSP_SpecialAtkName = "ゲージわざ_gsCdps"
Public Const SA1_CDSP_Value = "cDPS_gCdps"
Public Const SA1_CDSP_Cycle = "Cyc_gCdps"


Public Const SA1_CDST_NormalAtkName = "通常わざ_mnCdps"
Public Const SA1_CDST_SpecialAtkName = "ゲージわざ_msCdps"
Public Const SA1_CDST_Value = "cDPT_mCdps"

Public Const SA1_SA_Dpe = "tDPE最大" & C_SpecialAttack
Public Const SA1_SA_DpeAtkName = "ゲージtDPE技"
Public Const SA1_SA_DpeValue = "ゲージtDPE"

Public Const SA1_LastColumn As String = "ΔcDPT"

Public Const R_NormalAtkSpeciesSelect As String = "通常わざ種族選択"
Public Const R_SpecialAtkSpeciesSelect As String = "ゲージわざ種族選択"

'   種族マップ
Public Const R_SpeciesMapTypeSelect As String = "種族マップタイプ選択"
Public Const R_SpeciesMapSpeciesSelect As String = "種族マップ種族選択"
Public Const R_SpeciesMapSettings As String = "種族マップ設定"

'   ジム・対戦
Public Const BE_R_Settngs As String = "設定"
Public Const ME_R_Settngs As String = "対戦バトル設定"
Public Const BE_SetMode As String = "モード"
Public Const BE_SetSelfAtkDelay As String = "自攻撃遅延"
Public Const BE_SetEnemyAtkDelay As String = "敵攻撃遅延"
Public Const BE_SetRankNum As String = "順位数"
Public Const BE_SetRankVar As String = "順位付け"
Public Const BE_SetWithLimit As String = "限定技"
Public Const BE_DefCpUpper As String = "CP上限デフォルト"
Public Const BE_DefCpLower As String = "CP下限デフォルト"

Public Const BE_R_DummyEnemy As String = "ダミー"

Public Const BE_Species As String = C_SpeciesName
Public Const BE_Memo As String = "備考"
Public Const BE_NormalAttack As String = C_NormalAttack
Public Const BE_SpecialAttack As String = C_SpecialAttack
Public Const BE_SpecInput = "PL個体値"
Public Const BE_PL = "PL"
Public Const BE_ATK = "ATK"
Public Const BE_DEF = "DEF"
Public Const BE_HP = "HP"
Public Const BE_CPHP = "CP/HP"
'Public Const BE_HP = "HP"
Public Const BE_UpperCP = "上限"
Public Const BE_LowerCP = "下限"
Public Const BE_BaseRank = "天候未設定"
Public Const BE_SubRank = "天候設定"
Public Const BE_Weather As String = "天候"
Public Const BE_Rank As String = "順位"
Public Const BE_CtrName As String = "ニックネーム"
Public Const BE_CtrAttack As String = C_SpecialAttack
Public Const BE_KT As String = "KT"
Public Const BE_KTR As String = "KTR"
Public Const BE_RankOut As String = "ランク落ち"
Public Const BE_CalcTime As String = "計算時間"

Public Const BE_RankBase As String = "順位_b"

Public Const BE_SuffixBase As String = "_b"
Public Const BE_SuffixWeather As String = "_w"
Public Const BE_SuffixPredictBase As String = "_pb"
Public Const BE_SuffixPredictWeather As String = "_pw"

Public Const BE_NewEntryColorIndex As Integer = 30
Public Const BE_ReEntryColorIndex As Integer = 38

'   期待個体
Public Const NE_Name As String = "ニックネーム"
Public Const NE_EntryNum As String = "登場数"
Public Const NE_Type As String = "タイプ"
Public Const NE_PL As String = "PL"
Public Const NE_prPL As String = "PL_pr"
Public Const NE_Candies As String = "アメ"
Public Const NE_Sands As String = "星の砂"

Public Const NER_CountLower As String = "計数下限順位"
Public Const NELower_GymNow As String = "ジム現在"
Public Const NELower_GymPr As String = "ジム予測"
Public Const NELower_MtcNow As String = "対戦現在"
Public Const NELower_MtcPr As String = "対戦予測"

'   リスト
Public Const LI_R_Select As String = "リスト選択"
Public Const LI_Category As String = "カテゴリ"
Public Const LI_Note As String = "備考"
Public Const LI_Species As String = C_SpeciesName
Public Const LI_PL As String = "PL"
Public Const LI_ATK As String = "ATK"
Public Const LI_DEF As String = "DEF"
Public Const LI_HP As String = "HP"
Public Const LI_CP As String = "CP"

Public Const RCT_L0 As String = "サカキ"
Public Const RCT_L1 As String = "シエラ"
Public Const RCT_L2 As String = "グリフ"
Public Const RCT_L3 As String = "アルロ"

'   相関
Public Const R_Type As String = "タイプ"
Public Const R_InterTypeInflu As String = "相性表"
Public Const R_TypeSynonym As String = "タイプ略号"
Public Const R_WeatherBoost As String = "天候ブースト"
Public Const R_interTypeFactor As String = "タイプ相関係数"
Public Const R_WeatherFactor As String = "天候ブースト係数"
Public Const R_TypeMatchFactor As String = "タイプ一致係数"
Public Const R_MtcBtlFactor As String = "トレーナーバトル係数"
Public Const R_ChargeByDamage As String = "被ダメージによるチャージ"
Public Const R_WhiteSpace As String = "空白"
Public Const R_WeatherTable As String = "天候表"

'   CPM
Public Const R_StatusTransition As String = "わざの能力変化"
Public Const R_RocketTroupe As String = "ロケット団"
Public Const RT_Name As String = "名称"
Public Const RT_Number As String = "番目"
Public Const RT_Species As String = "種族名"

'   最大値
Public Const R_DummyParameter As String = "ダミーの敵のパラメータ"
Public Const DM_AtkPower As String = "攻撃力"
Public Const DM_DefPower As String = "防御力"
Public Const DM_HP As String = "HP"
Public Const DM_GymNAtkPower As String = "威力_ng"
Public Const DM_GymNAtkCharge As String = "チャージ_g"
Public Const DM_GymNAtkIdleTime As String = "発生時間"
Public Const DM_MtcNAtkPower As String = "威力_nm"
Public Const DM_MtcNAtkCharge As String = "チャージ_m"
Public Const DM_MtcNAtkIdleTurn As String = "発生ターン"
Public Const DM_GymSAtkPower As String = "威力_s"
Public Const DM_GymSAtkGuageNum As String = "ゲージ数"
Public Const DM_GymSAtkIdleTime As String = "発動時間_s"
Public Const DM_MtcSAtkPower As String = "威力_sm"
Public Const DM_MtcSAtkGuageVol As String = "ゲージ量"
'   Ver
Public Const VH_Branch As String = "ブランチ"
Public Const VH_Version As String = "Ver."
Public Const VH_Date As String = "日付"
Public Const VH_Article As String = "記事"

Public Const R_GlobalSettings As String = "全体設定"
Public Const GS_FileName As String = "ファイル名"
Public Const GS_BranchName As String = "ブランチ名"
Public Const GS_DistDir As String = "ディレクトリ"
Public Const GS_UseLimitedAttacksOnSpeciesAna As String = "種族分析で限定わざ"


'   メッセージ
Public Const msgConfirm As String = "本当に{0}してもいいですか？"
Public Const msgProcessing As String = "{1}を{0}しています。"
Public Const msgClear As String = "クリア"
Public Const msgRemove As String = "削除"
Public Const msgCalculate As String = "計算"
Public Const msgRanking As String = "ランキング"
Public Const msgExporting As String = "エクスポートしています。"
Public Const msgImporting As String = "インポートしています。"
Public Const msgAllSheet As String = "全てのシート"

Public Const msgMakingTable As String = "{0}テーブルを作成しています。"
Public Const msgMakingSheet As String = "{0}シートを作成しています。"
Public Const msgMaking As String = "{0}を作成しています。"
Public Const msgReseting As String = "{0}をリセットしています。"

Public Const msgSetColorToTypesOnTheSheet As String = "{0}シートのタイプに色を付けています。"
Public Const msgSetColorToTypesAndAttcksOnTheSheet As String = "{0}シートのタイプとわざ名に色を付けています。"


Public Const msgAddAttackToSpecies As String = "種族シートの{0}の{1}に{2}を追加しました。"
Public Const msgUnknownAttackName As String = "{1}が見つかりません。{0}シートに追加してください。"
Public Const msgKeyDoesNotExist As String = "{0}シートの{1}列に「{2}」がありません。"
Public Const msgColumnDoesNotExist As String = "{0}シートに「{1}」という列はありません。"
Public Const msgNoIdentifier As String = "{0}シートの{1}行{2}列に「{3}」がありません。"
Public Const msgColumnDoesNotExistOnTable As String = "「{0}」に「{1}」という列はありません。"

Public Const msgHitAttack As String = "{0}:{1}。ダメージ{2}。"
Public Const msgMonStatus As String = "{0}: HP:{1}, Rsv:{2}"

'   shSpecies
Public Const msgAttackIsLimited As String = "「{0}」を種族シートに追加します。「{0}」は限定わざですか？" _
                & vbCrLf & "[はい]=限定わざ、[いいえ]=普通のわざ、[キャンセル]=追加しない"
'   shNormalAttack, shSpecialAttack
Public Const msgResetSelectedAttack As String = "わざ名の選択をリセットしています。"
'   shSpeciesAnalysis1
Public Const msgSetNewSpeciesToSpeciesAnalysis1Sheet As String = "種族分析シートに種族を追加しています｡"
'   shIndividual
Public Const msgCalculatingIndividualSheet As String = "個体シートの解析値を計算しています。"
Public Const msgAligningIndividualSheet As String = "個体シートを調整しています。"
Public Const msgPLis0 As String = "PLが0の個体がありました。パラメータを確認してください。"
'   shIndivMap
Public Const msgAllRecalc As String = "全ての個体分析シートの再計算をしています。"
'   Ranking
Public Const msgCalcRank As String = "{0}の対策ランキングを計算しています。"
Public Const msgSetWildCard As String = "すべての相手にワイルドカードを設定しています。"
Public Const msgAddingRocketTroupe As String = "ロケット団を追加しています。"
'   Controls
Public Const msgSureToAllClear As String = "すべてクリアします。よろしいですか？"
Public Const msgDoesOpenLog As String = "ログがあります。開きますか？"
Public Const msgNoChange As String = "変更はありません。"
'   List
Public Const msgClearList As String = "リストをクリアしています。"




