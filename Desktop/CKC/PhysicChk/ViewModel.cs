using PhysicPlanCheck;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace PhysicPlanCheck
{
    public class ViewModel
    {
        public class GeneralInformation
        {
            public string PlanUid { get; set; }
            public string PatientName { get; set; }
            public string Course { get; set; }
            public string PlanName { get; set; }
            public string ApprovalStatus { get; set; }
            public string ImageName { get; set; }
            public string StructureSetName { get; set; }
            public string Scanner { get; set; }
            public string TreatmentUnit { get; set; }
            public string PlanType { get; set; }
            public string PlanTechnique { get; set; }
            //public string Machine { get; set; }
            public double TotalDose { get; set; }
            public double Fraction { get; set; }
            public bool IsHyperArc { get; set; }
            public string StructureOverride { get; set; }
            


            public object this[string propertyName]
            {
                get
                {
                    Type myType = typeof(GeneralInformation);
                    PropertyInfo myPropInfo = myType.GetProperty(propertyName);

                    return myPropInfo.GetValue(this);
                }

            }
        }
        public class CheckList
        {
            //Couch
            public bool BeamId;
            public bool BeamNumber;
            public bool? BolusAppliedToAllBeam;
            public bool? ControlPointMinimumMU;
            public bool DoseRate;
            public bool ExactIGRTCouchIsExist;
            public bool ExactIGRTCouchHeight;
            public bool HyperArcCouchIsExist;
            public bool? ImrtMuCalculation;
            public bool? JawFollowMlc;
            public bool? LeafClosedOutOfJaw;
            public bool? UserOriginIsOnCorrectPosition;
            public bool? BeamCenterIsOnCorrectPosition;
            public bool PrimaryReferencePointDoseLimit;
            public bool PrimaryReferencePointMatchPatientVolumeAndTargetVolumeId;
            public bool ReferencePointNotExistLocation;
            public bool TargetVolumeTypeIsCorrect;
            public bool TolerenceTableIsCorrect;
            public bool TreatmentUnitConsistent;
            public bool MultipleReferencePointNotExist;
            public bool? SetupFieldId;
            public bool SetupFieldIsExist;
            public bool? SetupFieldParameter;
        }
        //yitingEdit
        public string patientName { get; }
        public string patientBDay { get; }
        public string ImageDate { get; }

        public GeneralInformation[] generalinformation { get; set; }
        public CheckList[] PlanCheckList { get; set; }
        public static ScriptContext context { get; set; }
        public ViewModel(ScriptContext context_input)
        {
            context = context_input;

            {
                //yitingEdit
                patientName = context_input.Patient.Name.ToString();
                patientBDay = context_input.Patient.DateOfBirth.ToString();
                ImageDate = context_input.Image.CreationDateTime.ToString();

                List<PlanSetup> plan = new List<PlanSetup>();
                plan = CurrentPlan(plan);

                List<string> plantype = new List<string>();
                plantype = CheckPlanType(plan);

                List<string> plantechnique = new List<string>();
                plantechnique = EBRTCheckTechnique(plan);

                generalinformation = new GeneralInformation[plan.Count()];
                generalinformation = GetGeneralInformation(context, plan, plantype, plantechnique);

                PlanCheckList = new CheckList[plan.Count()];

                for (var i = 0; i < plan.Count(); i++)
                {
                    string msg = string.Empty;
                    CheckList tempPlanCheckList = new CheckList();

                    msg += string.Format("{0}\n{1},{2}\n{3}\n{4}\n{5},{6}\nIsHyperArc:{7}\nTotalDose:{8},Fraction:{9}\nTreatmentUnit: {10}\n==========================\n" +
                        "{11}", generalinformation[i].PatientName, generalinformation[i].Course,
                         generalinformation[i].PlanName, generalinformation[i].ImageName, generalinformation[i].Scanner,
                         generalinformation[i].PlanType, generalinformation[i].PlanTechnique, generalinformation[i].IsHyperArc,
                         generalinformation[i].TotalDose, generalinformation[i].Fraction, generalinformation[i].TreatmentUnit,
                         generalinformation[i].StructureOverride);

                    //msg += string.Format("\n");

                    //ShowAllPropertiesInGeneralInformation(generalinformation[i]);

                    tempPlanCheckList.ExactIGRTCouchIsExist = CheckIGRTCouch(plan[i]);
                    if (generalinformation[i].IsHyperArc) tempPlanCheckList.HyperArcCouchIsExist = CheckHyperArcCouch(plan[i]);
                    tempPlanCheckList.UserOriginIsOnCorrectPosition = CheckUserOrigin(plan[i]);
                    tempPlanCheckList.BeamCenterIsOnCorrectPosition = CheckBeamCenter(plan[i]);
                    tempPlanCheckList.TargetVolumeTypeIsCorrect = CheckTargetVolumeType(plan[i]);
                    tempPlanCheckList.TolerenceTableIsCorrect = CheckTolerenceTableForEBRT(plan[i], generalinformation[i].PlanTechnique);
                    tempPlanCheckList.PrimaryReferencePointDoseLimit = CheckPrimaryReferencePointDoseLimit(plan[i]);
                    tempPlanCheckList.PrimaryReferencePointMatchPatientVolumeAndTargetVolumeId = CheckPrimaryReferencePointMatchPatientVolumeAndTargetVolumeId(plan[i]);
                    tempPlanCheckList.MultipleReferencePointNotExist = CheckWhetherMultipleReferencePointNotExist(plan[i]);
                    tempPlanCheckList.ReferencePointNotExistLocation = CheckReferencePointNotExistLocation(plan[i]);
                    tempPlanCheckList.DoseRate = CheckDoseRateForEBRT(plan[i], generalinformation[i].PlanTechnique);
                    tempPlanCheckList.SetupFieldId = CheckSetupFieldId(plan[i]);
                    tempPlanCheckList.SetupFieldIsExist = CheckMandatorySetupFieldIsExist(plan[i]);
                    tempPlanCheckList.SetupFieldParameter = CheckSetupFieldParameters(plan[i]);
                    tempPlanCheckList.ControlPointMinimumMU = CheckControlPointMinimumMU(plan[i]);
                    tempPlanCheckList.BeamId = CheckBeamId(plan[i]);
                    tempPlanCheckList.TreatmentUnitConsistent = CheckTreatmentUnitConsistency(plan[i]);
                    tempPlanCheckList.LeafClosedOutOfJaw = CheckLeafClosedOutOfJaw(plan[i]);
                    tempPlanCheckList.BeamNumber = CheckBeamNumber(plan[i]);
                    tempPlanCheckList.JawFollowMlc = CheckJawFolloweMlc(plan[i]);
                    tempPlanCheckList.ImrtMuCalculation = CheckImrtMuCalculation(plan[i]);
                    tempPlanCheckList.BolusAppliedToAllBeam = CheckBolusAreAppliedToAllBeam(plan[i]);

                    PlanCheckList[i] = tempPlanCheckList;
                    /*
                    if (PlanCheckList[i].ExactIGRTCouchIsExist != true) msg += string.Format("IsIGRTCouch={0}\n", PlanCheckList[i].ExactIGRTCouchIsExist);
                    if (PlanCheckList[i].HyperArcCouchIsExist != true) msg += string.Format("IsHyperArcCouch={0}\n", PlanCheckList[i].HyperArcCouchIsExist);
                    if (PlanCheckList[i].UserOriginIsOnCorrectPosition != true) msg += string.Format("UserOrigin={0}\n", PlanCheckList[i].UserOriginIsOnCorrectPosition);
                    if (PlanCheckList[i].BeamCenterIsOnCorrectPosition != true) msg += string.Format("BeamCenter={0}\n", PlanCheckList[i].BeamCenterIsOnCorrectPosition);
                    if (PlanCheckList[i].TargetVolumeTypeIsCorrect != true) msg += string.Format("TargetVolumeType={0}\n", PlanCheckList[i].TargetVolumeTypeIsCorrect);
                    if (PlanCheckList[i].TolerenceTableIsCorrect != true) msg += string.Format("ToleranceTable={0}\n", PlanCheckList[i].TolerenceTableIsCorrect);
                    if (PlanCheckList[i].PrimaryReferencePointDoseLimit != true) msg += string.Format("PrimaryReferencePointDoseLimit={0}\n", PlanCheckList[i].PrimaryReferencePointDoseLimit);
                    if (PlanCheckList[i].PrimaryReferencePointMatchPatientVolumeAndTargetVolumeId != true) msg += string.Format("PrimaryReferencePointId={0}\n", PlanCheckList[i].PrimaryReferencePointMatchPatientVolumeAndTargetVolumeId);
                    if (PlanCheckList[i].MultipleReferencePointNotExist != true) msg += string.Format("MultipleReferencePointNotExcist={0}\n", PlanCheckList[i].MultipleReferencePointNotExist);
                    if (PlanCheckList[i].ReferencePointNotExistLocation != true) msg += string.Format("ReferencePointNotExistLocation={0}\n", PlanCheckList[i].ReferencePointNotExistLocation);
                    if (PlanCheckList[i].DoseRate != true) msg += string.Format("DoseRate={0}\n", PlanCheckList[i].DoseRate);
                    if (PlanCheckList[i].SetupFieldId != true) msg += string.Format("SetupFieldId={0}\n", PlanCheckList[i].SetupFieldId);
                    if (PlanCheckList[i].SetupFieldIsExist != true) msg += string.Format("SetupFieldIsExist={0}\n", PlanCheckList[i].SetupFieldIsExist);
                    if (PlanCheckList[i].SetupFieldParameter != true) msg += string.Format("SetupFieldParameter={0}\n", PlanCheckList[i].SetupFieldParameter);
                    if (PlanCheckList[i].ControlPointMinimumMU != true) msg += string.Format("ControlPointMinimumMU={0}\n", PlanCheckList[i].ControlPointMinimumMU);
                    if (PlanCheckList[i].BeamId != true) msg += string.Format("BeamId={0}\n", PlanCheckList[i].BeamId);
                    if (PlanCheckList[i].ImrtMuCalculation != true) msg += string.Format("IMRTMUCalculation={0}\n", PlanCheckList[i].ImrtMuCalculation);
                    if (PlanCheckList[i].BolusAppliedToAllBeam != true) msg += string.Format("BolusIsAppliedToAllBeam={0}\n", PlanCheckList[i].BolusAppliedToAllBeam);
                    if (PlanCheckList[i].TreatmentUnitConsistent != true) msg += string.Format("TreatmentUnitConsistent={0}\n", PlanCheckList[i].TreatmentUnitConsistent);
                    if (PlanCheckList[i].LeafClosedOutOfJaw != true) msg += string.Format("IsLeafClosedOutOfJaw={0}\n", PlanCheckList[i].LeafClosedOutOfJaw);
                    if (PlanCheckList[i].BeamNumber != true) msg += string.Format("IsBeamNumberCorrect={0}\n", PlanCheckList[i].BeamNumber);
                    if (PlanCheckList[i].JawFollowMlc != true) msg += string.Format("IsJawFollowMlc={0}\n", PlanCheckList[i].JawFollowMlc);
                    */
                    switch (generalinformation[i].PlanTechnique)
                    {
                        case "VMAT":
                            if (PlanCheckList[i].ExactIGRTCouchIsExist != true) msg += string.Format("IsIGRTCouch\n");
                            if (PlanCheckList[i].UserOriginIsOnCorrectPosition != true) msg += string.Format("UserOrigin\n");
                            if (PlanCheckList[i].BeamCenterIsOnCorrectPosition != true) msg += string.Format("BeamCenter\n");
                            if (PlanCheckList[i].TargetVolumeTypeIsCorrect != true) msg += string.Format("TargetVolumeType\n");
                            if (PlanCheckList[i].TolerenceTableIsCorrect != true) msg += string.Format("ToleranceTable\n");
                            if (PlanCheckList[i].PrimaryReferencePointDoseLimit != true) msg += string.Format("PrimaryReferencePointDoseLimit\n");
                            if (PlanCheckList[i].PrimaryReferencePointMatchPatientVolumeAndTargetVolumeId != true) msg += string.Format("PrimaryReferencePointId\n");
                            if (PlanCheckList[i].MultipleReferencePointNotExist != true) msg += string.Format("MultipleReferencePointNotExcist\n");
                            if (PlanCheckList[i].ReferencePointNotExistLocation != true) msg += string.Format("ReferencePointNotExistLocation\n");
                            if (PlanCheckList[i].DoseRate != true) msg += string.Format("DoseRate\n");
                            if (PlanCheckList[i].SetupFieldId != true) msg += string.Format("SetupFieldId\n");
                            if (PlanCheckList[i].SetupFieldIsExist != true) msg += string.Format("SetupFieldIsExist\n");
                            if (PlanCheckList[i].SetupFieldParameter != true) msg += string.Format("SetupFieldParameter\n");
                            if (PlanCheckList[i].BeamId != true) msg += string.Format("BeamId\n");
                            if (PlanCheckList[i].BolusAppliedToAllBeam != true) msg += string.Format("BolusIsAppliedToAllBeam\n");
                            if (PlanCheckList[i].TreatmentUnitConsistent != true) msg += string.Format("TreatmentUnitConsistent\n");
                            if (PlanCheckList[i].BeamNumber != true) msg += string.Format("IsBeamNumberCorrect\n");
                            break;

                        case "IMRT":
                            if (PlanCheckList[i].ExactIGRTCouchIsExist != true) msg += string.Format("IsIGRTCouch\n");
                            if (PlanCheckList[i].UserOriginIsOnCorrectPosition != true) msg += string.Format("UserOrigin\n");
                            if (PlanCheckList[i].BeamCenterIsOnCorrectPosition != true) msg += string.Format("BeamCenter\n");
                            if (PlanCheckList[i].TargetVolumeTypeIsCorrect != true) msg += string.Format("TargetVolumeType\n");
                            if (PlanCheckList[i].TolerenceTableIsCorrect != true) msg += string.Format("ToleranceTable\n");
                            if (PlanCheckList[i].PrimaryReferencePointDoseLimit != true) msg += string.Format("PrimaryReferencePointDoseLimit\n");
                            if (PlanCheckList[i].PrimaryReferencePointMatchPatientVolumeAndTargetVolumeId != true) msg += string.Format("PrimaryReferencePointId\n");
                            if (PlanCheckList[i].MultipleReferencePointNotExist != true) msg += string.Format("MultipleReferencePointNotExcist\n");
                            if (PlanCheckList[i].ReferencePointNotExistLocation != true) msg += string.Format("ReferencePointNotExistLocation\n");
                            if (PlanCheckList[i].DoseRate != true) msg += string.Format("DoseRate\n");
                            if (PlanCheckList[i].SetupFieldId != true) msg += string.Format("SetupFieldId\n");
                            if (PlanCheckList[i].SetupFieldIsExist != true) msg += string.Format("SetupFieldIsExist\n");
                            if (PlanCheckList[i].SetupFieldParameter != true) msg += string.Format("SetupFieldParameter\n");
                            if (PlanCheckList[i].BeamId != true) msg += string.Format("BeamId\n");
                            if (PlanCheckList[i].ImrtMuCalculation != true) msg += string.Format("IMRTMUCalculation\n");
                            if (PlanCheckList[i].BolusAppliedToAllBeam != true) msg += string.Format("BolusIsAppliedToAllBeam\n");
                            if (PlanCheckList[i].TreatmentUnitConsistent != true) msg += string.Format("TreatmentUnitConsistent\n");
                            if (PlanCheckList[i].BeamNumber != true) msg += string.Format("IsBeamNumberCorrect\n");
                            break;

                        case "FiF":
                        case "3or2D":
                            if (PlanCheckList[i].ExactIGRTCouchIsExist != true) msg += string.Format("IsIGRTCouch\n");
                            if (PlanCheckList[i].UserOriginIsOnCorrectPosition != true) msg += string.Format("UserOrigin\n");
                            if (PlanCheckList[i].BeamCenterIsOnCorrectPosition != true) msg += string.Format("BeamCenter\n");
                            if (PlanCheckList[i].TargetVolumeTypeIsCorrect != true) msg += string.Format("TargetVolumeType\n");
                            if (PlanCheckList[i].TolerenceTableIsCorrect != true) msg += string.Format("ToleranceTable\n");
                            if (PlanCheckList[i].PrimaryReferencePointDoseLimit != true) msg += string.Format("PrimaryReferencePointDoseLimit\n");
                            if (PlanCheckList[i].PrimaryReferencePointMatchPatientVolumeAndTargetVolumeId != true) msg += string.Format("PrimaryReferencePointId\n");
                            if (PlanCheckList[i].MultipleReferencePointNotExist != true) msg += string.Format("MultipleReferencePointNotExcist\n");
                            if (PlanCheckList[i].ReferencePointNotExistLocation != true) msg += string.Format("ReferencePointNotExistLocation\n");
                            if (PlanCheckList[i].DoseRate != true) msg += string.Format("DoseRate\n");
                            if (PlanCheckList[i].SetupFieldId != true) msg += string.Format("SetupFieldId\n");
                            if (PlanCheckList[i].SetupFieldParameter != true) msg += string.Format("SetupFieldParamete}\n");
                            if (PlanCheckList[i].ControlPointMinimumMU != true) msg += string.Format("ControlPointMinimumMU\n");
                            if (PlanCheckList[i].BeamId != true) msg += string.Format("BeamId\n");
                            if (PlanCheckList[i].BolusAppliedToAllBeam != true) msg += string.Format("BolusIsAppliedToAllBeam\n");
                            if (PlanCheckList[i].TreatmentUnitConsistent != true) msg += string.Format("TreatmentUnitConsistent\n");
                            if (PlanCheckList[i].LeafClosedOutOfJaw != true) msg += string.Format("IsLeafClosedOutOfJaw\n");
                            if (PlanCheckList[i].BeamNumber != true) msg += string.Format("IsBeamNumberCorrect\n");
                            //if (PlanCheckList[i].JawFollowMlc != true) msg += string.Format("IsJawFollowMlc\n");
                            break;

                        case "3or2DwithoutMLC":
                            if (PlanCheckList[i].ExactIGRTCouchIsExist != true) msg += string.Format("IsIGRTCouch={0}\n");
                            if (PlanCheckList[i].UserOriginIsOnCorrectPosition != true) msg += string.Format("UserOrigin={0}\n");
                            if (PlanCheckList[i].BeamCenterIsOnCorrectPosition != true) msg += string.Format("BeamCenter={0}\n");
                            if (PlanCheckList[i].TargetVolumeTypeIsCorrect != true) msg += string.Format("TargetVolumeType={0}\n");
                            if (PlanCheckList[i].TolerenceTableIsCorrect != true) msg += string.Format("ToleranceTable={0}\n");
                            if (PlanCheckList[i].PrimaryReferencePointDoseLimit != true) msg += string.Format("PrimaryReferencePointDoseLimit={0}\n");
                            if (PlanCheckList[i].PrimaryReferencePointMatchPatientVolumeAndTargetVolumeId != true) msg += string.Format("PrimaryReferencePointId={0}\n");
                            if (PlanCheckList[i].MultipleReferencePointNotExist != true) msg += string.Format("MultipleReferencePointNotExcist={0}\n");
                            if (PlanCheckList[i].ReferencePointNotExistLocation != true) msg += string.Format("ReferencePointNotExistLocation={0}\n");
                            if (PlanCheckList[i].DoseRate != true) msg += string.Format("DoseRate={0}\n");
                            if (PlanCheckList[i].SetupFieldId != true) msg += string.Format("SetupFieldId={0}\n");
                            if (PlanCheckList[i].SetupFieldParameter != true) msg += string.Format("SetupFieldParameter={0}\n");
                            if (PlanCheckList[i].ControlPointMinimumMU != true) msg += string.Format("ControlPointMinimumMU={0}\n");
                            if (PlanCheckList[i].BeamId != true) msg += string.Format("BeamId={0}\n");
                            if (PlanCheckList[i].BolusAppliedToAllBeam != true) msg += string.Format("BolusIsAppliedToAllBeam={0}\n");
                            if (PlanCheckList[i].TreatmentUnitConsistent != true) msg += string.Format("TreatmentUnitConsistent={0}\n");
                            if (PlanCheckList[i].BeamNumber != true) msg += string.Format("IsBeamNumberCorrect={0}\n");
                            break;
                    }


                    MessageBox.Show(msg);
                    //OutputResult(msg, generalinformation[i]);

                }

            }

        }

        public void ShowAllPropertiesInGeneralInformation(GeneralInformation generalinformation)
        {
            string msg = string.Empty;
            PropertyInfo[] generalinformation_member = generalinformation.GetType().GetProperties();

            //j只到generalinformation property數量的-1，目的是避開最後一個用於回報property value的method。
            for (var j = 0; j < generalinformation_member.Count() - 1; j++)
            {
                msg += string.Format("{0},{1}\n", generalinformation_member[j].Name, generalinformation[generalinformation_member[j].Name]);
            }
            MessageBox.Show(msg);
        }

        public class StructureWithDensity
        {
            public string StructureId;
            public double HUValue;
        }

        public class MLCparameters
        {
            public int NumberInner { get; private set; }
            public int NumberOutter { get; private set; }
            public double ThicknessInner { get; private set; }
            public double ThicknessOutter { get; private set; }
            public int NumberLeafPair { get; private set; }

            public MLCparameters(string Treatmentunit)
            {
                switch (Treatmentunit)
                {
                    case "LA5TB2069":
                    case "LA7TB4557":
                        this.NumberInner = 16;
                        this.NumberOutter = 14;
                        this.ThicknessInner = 0.25;
                        this.ThicknessOutter = 0.5;
                        this.NumberLeafPair = 60;
                        break;

                    case "LA6TB4313":
                    case "LA3TB1623":
                        this.NumberInner = 20;
                        this.NumberOutter = 10;
                        this.ThicknessInner = 0.5;
                        this.ThicknessOutter = 1;
                        this.NumberLeafPair = 60;
                        break;
                    default:
                        this.NumberInner = 20;
                        this.NumberOutter = 10;
                        this.ThicknessInner = 0.5;
                        this.ThicknessOutter = 1;
                        this.NumberLeafPair = 60;
                        break;

                }
            }
        }

        //Check Whether jaw follow MLC
        public bool? CheckJawFolloweMlc(PlanSetup plan)
        {
            bool? IsJawFollowMlc = true;
            double Criteria = 0.1, CriteriaJawOverTravel = 0.05;//0.1cm接受Jaw與Leaf不match最大0.1cm, Jaw關超過Leaf 0.05cm
            IEnumerable<Beam> TreatmentBeam = plan.Beams.Where(x => x.IsSetupField == false);
            List<PlanSetup> dummyPlanList = new List<PlanSetup>();
            dummyPlanList.Add(plan);
            List<string> PlanTechnique = EBRTCheckTechnique(dummyPlanList);
            double JawX1 = 0.0, JawX2 = 0.0;
            double OutterLeafX1 = 1000000.0, OutterLeafX2 = -1000000.0;
            double DifferenceX1 = 0.0, DifferenceX2 = 0.0;
            int RecordLeafNumberX1 = 0, RecordLeafNumberX2 = 0;
            int index = 0, j = 0;
            string msg = string.Empty;
            string Treatmentunit = TreatmentBeam.First().TreatmentUnit.Id;
            List<int[]> LeafOpeningList = DeterminLeafShouldOpeningByJaw(Treatmentunit, plan);//根據Jaw的位置找出應該打開與應該關閉的leaf

            if (string.Equals(PlanTechnique[0], "FiF") || string.Equals(PlanTechnique[0], "3or2D"))
            {
                foreach (var beam in TreatmentBeam)
                {
                    OutterLeafX1 = 1000000.0;
                    OutterLeafX2 = -1000000.0;

                    List<ControlPoint> ControlPoint = beam.ControlPoints.ToList();

                    JawX2 = ControlPoint[0].JawPositions.X2 / 10;
                    JawX1 = ControlPoint[0].JawPositions.X1 / 10;

                    //msg += string.Format("{0},{1}\n", LeafOpeningList[index][1], LeafOpeningList[index][0]);
                    for (j = LeafOpeningList[index][1]; j <= LeafOpeningList[index][0]; j++)
                    {
                        if (ControlPoint[0].LeafPositions[1, j - 1] / 10 > OutterLeafX2)
                        {
                            OutterLeafX2 = ControlPoint[0].LeafPositions[1, j - 1] / 10;
                            RecordLeafNumberX2 = j;
                        }

                        if (ControlPoint[0].LeafPositions[0, j - 1] / 10 < OutterLeafX1)
                        {
                            OutterLeafX1 = ControlPoint[0].LeafPositions[0, j - 1] / 10;
                            RecordLeafNumberX1 = j;
                        }
                    }

                    DifferenceX2 = Math.Round(JawX2 - OutterLeafX2, 3);
                    DifferenceX1 = Math.Round(JawX1 - OutterLeafX1, 3);
                    //msg += string.Format("{0},{1}\n", DifferenceX1, JawX1);
                    if (Math.Abs(DifferenceX2) < Criteria && DifferenceX2 >= -CriteriaJawOverTravel)
                    {

                    }
                    else
                    {
                        msg += string.Format("Beam:{0}, BankB Leaf:{1}, Jaw-Leaf= {2}\n", beam.Id, RecordLeafNumberX2, DifferenceX2);
                        IsJawFollowMlc = false;
                    }

                    if (Math.Abs(DifferenceX1) < Criteria && DifferenceX1 <= CriteriaJawOverTravel)
                    {

                    }
                    else
                    {
                        msg += string.Format("Beam:{0}, BankA Leaf:{1}, Jaw-Leaf= {2}\n", beam.Id, RecordLeafNumberX1, DifferenceX1);
                        IsJawFollowMlc = false;
                    }
                    index++;
                }
            }
            else
            {
                IsJawFollowMlc = null;
            }

            if (string.IsNullOrEmpty(msg) != true)
            {
                MessageBox.Show(msg, "Waring: Jaw not follow Leaf");
            }
            return IsJawFollowMlc;
        }

        //根據Jaw的位置回傳每一隻beam 中 leaf需要打開的根數[開始打開, 最後一根打開]。此回傳結果為每一隻beam最後一個control point的結果，
        //其假設beam裡面所有control point Y Jaw的位置都是相同的。
        //[Y2,Y1]: 1-60
        public List<int[]> DeterminLeafShouldOpeningByJaw(string Treatmentunit, PlanSetup plan)
        {
            IEnumerable<Beam> TreatmentBeam = plan.Beams.Where(x => x.IsSetupField == false);
            MLCparameters MlcParameter = new MLCparameters(Treatmentunit);
            double JawY1 = 0.0, JawY2 = 0.0;
            int i = 0;
            int FirstOpeningLeaf = 0, LastOpeningLeaf = 0;
            int[] LeafOpeningArray;
            List<int[]> LeafOpeningList = new List<int[]>();
            string msg = string.Empty;

            List<PlanSetup> dummyPlanList = new List<PlanSetup>();
            dummyPlanList.Add(plan);
            List<string> PlanTechnique = EBRTCheckTechnique(dummyPlanList);
            bool IsFiFor3Dor2D = false;
            if (string.Equals(PlanTechnique[0], "FiF") || string.Equals(PlanTechnique[0], "3or2D"))
            {
                IsFiFor3Dor2D = true;
            }

            if (IsFiFor3Dor2D == true)
            {
                foreach (var beam in TreatmentBeam)
                {
                    if (beam.MLCPlanType.ToString() == "Static" || beam.MLCPlanType.ToString() == "DoseDynamic")
                    {
                        List<ControlPoint> ControlPoint = beam.ControlPoints.ToList();
                        for (i = 0; i < ControlPoint.Count(); i = i + 2)
                        {
                            JawY2 = beam.ControlPoints[i].JawPositions.Y2 / 10;
                            JawY1 = beam.ControlPoints[i].JawPositions.Y1 / 10;

                            if (JawY2 >= 0)
                            {
                                if (JawY2 > MlcParameter.ThicknessInner * MlcParameter.NumberInner)
                                {
                                    FirstOpeningLeaf = Convert.ToInt32(MlcParameter.NumberInner + MlcParameter.NumberOutter + MlcParameter.NumberInner +
                                        Math.Ceiling((JawY2 - MlcParameter.NumberInner * MlcParameter.ThicknessInner) / MlcParameter.ThicknessOutter));
                                }
                                else
                                {
                                    FirstOpeningLeaf = Convert.ToInt32(MlcParameter.NumberInner + MlcParameter.NumberOutter + Math.Ceiling(JawY2 / MlcParameter.ThicknessInner));
                                }
                            }
                            else
                            {
                                if (JawY2 <= -MlcParameter.ThicknessInner * MlcParameter.NumberInner)
                                {
                                    FirstOpeningLeaf = Convert.ToInt32(MlcParameter.NumberInner + MlcParameter.NumberOutter + 1 - MlcParameter.NumberInner -
                                        Math.Ceiling((Math.Abs(JawY2) - MlcParameter.NumberInner * MlcParameter.ThicknessInner) / MlcParameter.ThicknessOutter));
                                }
                                else
                                {
                                    FirstOpeningLeaf = Convert.ToInt32(MlcParameter.NumberInner + MlcParameter.NumberOutter + 1 - Math.Ceiling(Math.Abs(JawY2) /
                                        MlcParameter.ThicknessInner));
                                }
                            }

                            if (JawY1 <= 0)
                            {
                                if (JawY1 <= -MlcParameter.ThicknessInner * MlcParameter.NumberInner)
                                {
                                    LastOpeningLeaf = Convert.ToInt32(MlcParameter.NumberInner + MlcParameter.NumberOutter + 1 - MlcParameter.NumberInner -
                                        Math.Ceiling((Math.Abs(JawY1) - MlcParameter.NumberInner * MlcParameter.ThicknessInner) / MlcParameter.ThicknessOutter));
                                }
                                else
                                {
                                    LastOpeningLeaf = Convert.ToInt32(MlcParameter.NumberInner + MlcParameter.NumberOutter + 1 - Math.Ceiling(Math.Abs(JawY1) /
                                        MlcParameter.ThicknessInner));
                                }
                            }
                            else
                            {
                                if (JawY1 > MlcParameter.ThicknessInner * MlcParameter.NumberInner)
                                {
                                    LastOpeningLeaf = Convert.ToInt32(MlcParameter.NumberInner + MlcParameter.NumberOutter + MlcParameter.NumberInner +
                                        Math.Ceiling((JawY1 - MlcParameter.NumberInner * MlcParameter.ThicknessInner) / MlcParameter.ThicknessOutter));
                                }
                                else
                                {
                                    LastOpeningLeaf = Convert.ToInt32(MlcParameter.NumberInner + MlcParameter.NumberOutter + Math.Ceiling(JawY1 / MlcParameter.ThicknessInner));
                                }
                            }
                            //msg += string.Format("{0},{1}\n", FirstOpeningLeaf,LastOpeningLeaf);                        
                        }
                        //MessageBox.Show(msg);
                    }
                    LeafOpeningArray = new int[] { FirstOpeningLeaf, LastOpeningLeaf };
                    LeafOpeningList.Add(LeafOpeningArray);
                }
            }
            return LeafOpeningList;
        }


        //Check leaf are all closed out of jaw
        public bool? CheckLeafClosedOutOfJaw(PlanSetup plan)
        {
            bool? IsLeafClosedOutOfJaw = false;
            IEnumerable<Beam> TreatmentBeam = plan.Beams.Where(x => x.IsSetupField == false);
            string Treatmentunit = TreatmentBeam.First().TreatmentUnit.Id;
            MLCparameters MlcParameter = new MLCparameters(Treatmentunit);
            List<int[]> LeafOpeningList = DeterminLeafShouldOpeningByJaw(Treatmentunit, plan);//根據Jaw的位置找出應該打開與應該關閉的leaf
            int index = 0, i = 0, j = 0;
            Single[,] LeafPosition;

            string msg = string.Empty;
            if (LeafOpeningList.Any())
            {
                foreach (var beam in TreatmentBeam)
                {
                    List<ControlPoint> ControlPoints = beam.ControlPoints.ToList();
                    for (i = 0; i < ControlPoints.Count(); i = i + 2)
                    {
                        LeafPosition = ControlPoints[i].LeafPositions;
                        for (j = 0; j < MlcParameter.NumberLeafPair; j++)
                        {
                            if (j + 1 > LeafOpeningList[index][0] || j + 1 < LeafOpeningList[index][1])
                            {
                                if (LeafPosition[0, j] != LeafPosition[1, j]) msg += string.Format("Beam:{0},CP:{1},Leaf:{2}\n", beam.Id, i, j + 1);
                            }
                        }
                    }
                    index++;
                }
            }
            else { IsLeafClosedOutOfJaw = null; }

            if (IsLeafClosedOutOfJaw == null)
            {
                return IsLeafClosedOutOfJaw;
            }
            else if (string.IsNullOrEmpty(msg))
            {
                IsLeafClosedOutOfJaw = true;
                return IsLeafClosedOutOfJaw;
            }
            else
            {
                IsLeafClosedOutOfJaw = false;
                MessageBox.Show(msg, "Warning: Leaf isn't closed");
                return IsLeafClosedOutOfJaw;

            }

        }

        public bool CheckTreatmentUnitConsistency(PlanSetup plan)
        {
            bool WhetherTreatmentUnitConsistent = false;
            IEnumerable<Beam> TreatmentBeam = plan.Beams.Where(x => x.IsSetupField == false);
            string Treatmentunit = TreatmentBeam.First().TreatmentUnit.Id;
            IEnumerable<Beam> TreatmentUnitNotConsistent = TreatmentBeam.Where(x => string.Equals(x.TreatmentUnit.Id, Treatmentunit) == false);

            if (TreatmentUnitNotConsistent.Any())
            {
                WhetherTreatmentUnitConsistent = false;
            }
            else
            {
                WhetherTreatmentUnitConsistent = true;
            }

            return WhetherTreatmentUnitConsistent;
        }

        //Check Bolus are applied to all beam
        public bool? CheckBolusAreAppliedToAllBeam(PlanSetup plan)
        {
            bool? IsBolusAppliedToAllBeam = false;
            IEnumerable<Beam> TreatmentBeam = plan.Beams.Where(x => x.IsSetupField == false);
            IEnumerable<Beam> BolusAppliedBeam = TreatmentBeam.Where(x => x.Boluses.Any() == true);
            IEnumerable<Structure> BolusStructure = plan.StructureSet.Structures.Where(x => x.Id.IndexOf("Bolus") != -1);
            int TreatmentBeamNumber = 0, BolusAppliedBeamNumber = 0;

            TreatmentBeamNumber = TreatmentBeam.Count();
            BolusAppliedBeamNumber = BolusAppliedBeam.Count();

            string msg = string.Empty;
            foreach (var i in BolusAppliedBeam)
            {
                msg += string.Format("{0}\n", i.Id);
            }

            if (BolusAppliedBeamNumber != 0)
            {
                string BolusMessage = string.Format("{0}:\n{1}", "Please Check Bolus is applied to the following beams", msg);
                MessageBox.Show(BolusMessage);
            }

            if (BolusAppliedBeamNumber == 0 && BolusStructure.Count() == 0)
            {
                IsBolusAppliedToAllBeam = true;
            }
            else if (TreatmentBeamNumber == BolusAppliedBeamNumber)
            {
                IsBolusAppliedToAllBeam = true;
            }
            else
            {
                IsBolusAppliedToAllBeam = false;
            }

            return IsBolusAppliedToAllBeam;
        }

        //Check IMRT MU Calculation 若有差異大於2MU的，則回傳Beam名稱。若回傳"NOT IMRT"表示該plan非IMRT Plan
        public bool? CheckImrtMuCalculation(PlanSetup plan)
        {
            bool? IsMUDifferent = false;
            IEnumerable<Beam> TreatmentBeam = plan.Beams.Where(x => x.IsSetupField == false);
            string Identify_LostMUFactor = "LostMUFactor", Identify_MaximumMU = "Maximum MUs";
            double LostMUFactor = 0.0, MaximumMU = 0.0, BeamMU = 0.0, CalculationMU = 0.0;
            double Criteria = 2.0;
            List<string> FailedBeam = new List<string>();
            List<PlanSetup> dummyPlanList = new List<PlanSetup>();
            dummyPlanList.Add(plan);
            List<string> PlanTechnique = EBRTCheckTechnique(dummyPlanList);

            string msg = string.Empty;
            if (string.Equals(PlanTechnique[0], "IMRT"))
            {
                foreach (var beam in TreatmentBeam)
                {
                    foreach (var log in beam.CalculationLogs)
                        foreach (var information in log.MessageLines)
                        {
                            if (information.IndexOf(Identify_LostMUFactor) != -1)
                            {
                                LostMUFactor = ExtractGantryAngleFromFieldId(information, " ", information.IndexOf(Identify_LostMUFactor) + Identify_LostMUFactor.Count());
                            }
                            else if (information.IndexOf(Identify_MaximumMU) != -1)
                            {
                                MaximumMU = ExtractGantryAngleFromFieldId(information, " ", information.IndexOf(Identify_MaximumMU) + Identify_MaximumMU.Count());
                            }
                        }

                    BeamMU = Math.Round(beam.Meterset.Value, 2);
                    CalculationMU = Math.Round(LostMUFactor * MaximumMU, 2);

                    if (Math.Abs(BeamMU - CalculationMU) >= Criteria)
                    {
                        FailedBeam.Add(beam.Id);
                    }
                    msg += string.Format("{0},BeamMU = {1},CalculationMU = {2}\n", beam.Id, BeamMU, CalculationMU);
                }

                IsMUDifferent = !FailedBeam.Any();
                if (IsMUDifferent == false)
                {
                    string Caption = string.Format("MU Difference > {0}MU", Criteria);
                    MessageBox.Show(msg, Caption);
                }
            }
            else
            {
                IsMUDifferent = null;
            }
            return IsMUDifferent;
        }

        //Check Control Point Minimum MU
        public bool? CheckControlPointMinimumMU(PlanSetup plan)
        {
            bool? IsAllBeamGreaterThanCriteria = true;
            double CriteriaMinimumMU = 4.0, ControlPointMU = 0.0;
            int j = 0;
            List<PlanSetup> InputPlanList = new List<PlanSetup> { plan }; //input plan is single plan
            List<string> OutputList = EBRTCheckTechnique(InputPlanList);
            string PlanTechnique = OutputList[0];
            List<double> ControlPointMeterSetWeight;
            IEnumerable<Beam> TreatmentBeam = plan.Beams.Where(x => x.IsSetupField == false);

            if (Equals(PlanTechnique, "FiF"))
            {
                foreach (var i in TreatmentBeam)
                {
                    ControlPointMeterSetWeight = i.ControlPoints.Select(x => x.MetersetWeight).ToList();
                    for (j = 0; j < ControlPointMeterSetWeight.Count(); j = j + 2)
                    {
                        ControlPointMU = (ControlPointMeterSetWeight[j + 1] - ControlPointMeterSetWeight[j]) * i.Meterset.Value;
                        if (ControlPointMU < CriteriaMinimumMU)
                        {
                            IsAllBeamGreaterThanCriteria = false;
                        }
                    }
                }

            }
            else if (Equals(PlanTechnique, "3or2D"))
            {
                IEnumerable<bool> BeamMUGreaterThanCriteria = TreatmentBeam.Select(x => x.Meterset.Value >= CriteriaMinimumMU);
                IsAllBeamGreaterThanCriteria = !BeamMUGreaterThanCriteria.Where(x => x = false).Any();
            }
            else
            {
                IsAllBeamGreaterThanCriteria = null;
            }

            return IsAllBeamGreaterThanCriteria;
        }

        //Find Number in String (要擷取數字的文字，辨識文字從後方開始抓數字，起始位數)
        public double ExtractGantryAngleFromFieldId(string InputString, string IdentifyCharater, int StartDigit)
        {
            int IdentifyCharacterLocation = (StartDigit > -1) ? InputString.IndexOf(IdentifyCharater, StartDigit) : -1;
            int i = 0;
            char SingleChar, ExceptionCharacter = Convert.ToChar(".");
            string FieldGantryAngle = string.Empty;
            double ReturnGantryAngle;

            //string msg = string.Empty;
            if (IdentifyCharacterLocation != -1)
            {

                while (IdentifyCharacterLocation + i + 1 < InputString.Length)
                {
                    i++;
                    SingleChar = InputString[IdentifyCharacterLocation + i];

                    if (Char.IsDigit(SingleChar))
                    {
                        FieldGantryAngle += string.Format("{0}", SingleChar);

                    }
                    else if (Equals(SingleChar, ExceptionCharacter))
                    {
                        FieldGantryAngle += string.Format("{0}", SingleChar);
                        //string msg = string.Empty;
                        //msg += string.Format("{0}\n", FieldGantryAngle);
                        //MessageBox.Show(msg);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            if (string.IsNullOrEmpty(FieldGantryAngle))
            {
                ReturnGantryAngle = Double.NaN;
            }
            else
            {
                ReturnGantryAngle = Convert.ToDouble(FieldGantryAngle);
            }
            return ReturnGantryAngle;
        }

        //Check Treatment Field Number
        public bool CheckBeamNumber(PlanSetup plan)
        {
            bool IsBeamNumberCorrect = true;
            IEnumerable<Beam> TreatmentField = plan.Beams.Where(x => x.IsSetupField == false);
            IEnumerable<double> TreatmentFieldNumber = TreatmentField.Select(x => ExtractGantryAngleFromFieldId(x.Id, "-", 0));
            IEnumerable<double> dummy;
            int i = 1;

            for (i = 1; i <= TreatmentField.Count(); i++)
            {
                dummy = TreatmentFieldNumber.Where(x => x == i);
                if (dummy.Any() == true && dummy.Count() < 2)
                {

                }
                else
                {
                    IsBeamNumberCorrect = false;
                    break;
                }
            }

            return IsBeamNumberCorrect;
            /*
            string msg = string.Empty;
            foreach (var i in TreatmentFieldNumber)
            {
                msg += string.Format("{0}", i);
            }
            MessageBox.Show(msg);*/
        }

        //Check Treatment Field ID
        public bool CheckBeamId(PlanSetup plan)
        {
            bool IsBeamIdAngleMatchPlan = true;
            IEnumerable<Beam> TreatmentField = plan.Beams.Where(x => x.IsSetupField == false);
            IEnumerable<double> PlanGantryAngleFirst = TreatmentField.Select(x => x.ControlPoints.First().GantryAngle);
            IEnumerable<double> PlanGantryAngleLast = TreatmentField.Select(x => x.ControlPoints.Last().GantryAngle);
            string BeamIdentifyCharacterFirst = "G", BeamIdentifyCharacterLast = "-";
            IEnumerable<double> BeamIdFirst = TreatmentField.Select(x => ExtractGantryAngleFromFieldId(x.Id, BeamIdentifyCharacterFirst, 0));
            IEnumerable<double> BeamIdLast = TreatmentField.Select(x => ExtractGantryAngleFromFieldId(x.Id, BeamIdentifyCharacterLast, x.Id.IndexOf(BeamIdentifyCharacterFirst)));

            List<double> PlanGantryAngle_A, PlanGantryAngle_B, BeamIdAngle_A, BeamIdAngle_B;

            PlanGantryAngle_A = PlanGantryAngleFirst.ToList();
            PlanGantryAngle_B = PlanGantryAngleLast.ToList();
            BeamIdAngle_A = BeamIdFirst.ToList();
            BeamIdAngle_B = BeamIdLast.ToList();

            int i = 0;
            //string msg = string.Empty;
            for (i = 0; i < TreatmentField.Count(); i++)
            {
                if (Double.IsNaN(BeamIdAngle_B[i]) != true)
                {
                    if (PlanGantryAngle_A[i] != BeamIdAngle_A[i] || PlanGantryAngle_B[i] != BeamIdAngle_B[i]) IsBeamIdAngleMatchPlan = false;
                }
                else
                {
                    if (PlanGantryAngle_A[i] != BeamIdAngle_A[i]) IsBeamIdAngleMatchPlan = false;
                }
                //msg += string.Format("{0},{1},{2},{3}\n",PlanGantryAngle_A[i], PlanGantryAngle_B[i], BeamIdAngle_A[i], BeamIdAngle_B[i]);
            }
            //MessageBox.Show(msg);
            return IsBeamIdAngleMatchPlan;
        }

        //Check SetupField ID
        public bool? CheckSetupFieldId(PlanSetup plan)
        {
            IEnumerable<Beam> SetupField = plan.Beams.Where(x => x.IsSetupField == true);
            double PlanGantryAngle = 0.0, FieldIdGantryAngle = 0.0;
            string SetupFieldId = string.Empty;
            string SetupIdentifyCharacter = "G", CBCTIdentifyCharacter = "CBCT";
            bool? IsSetupFieldIdCorrect = true;

            foreach (var i in SetupField)
            {
                PlanGantryAngle = i.ControlPoints.First().GantryAngle;
                SetupFieldId = i.Id;

                FieldIdGantryAngle = ExtractGantryAngleFromFieldId(SetupFieldId, SetupIdentifyCharacter, 0);

                if (PlanGantryAngle == FieldIdGantryAngle)
                {

                }
                else if (SetupFieldId.IndexOf(CBCTIdentifyCharacter) != -1)
                {

                }
                else
                {
                    IsSetupFieldIdCorrect = false;
                }

            }
            /*
            if (!SetupField.Any())
            {
                IsSetupFieldIdCorrect = null;
            }
            */
            return IsSetupFieldIdCorrect;
        }

        //Check Setup Field is Existed
        public bool CheckMandatorySetupFieldIsExist(PlanSetup plan)
        {
            bool SetupFieldIsExist = true;
            List<string> MandatorySetupField = new List<string> { "G0", "G90", "CBCT" };
            IEnumerable<Beam> SetupField = plan.Beams.Where(x => x.IsSetupField == true);
            IEnumerable<bool> dummy;
            VVector BeamIsocenter = new VVector();
            VVector BodyCenterPoint = plan.StructureSet.Structures.Where(x => Equals(x.DicomType, "EXTERNAL") == true).First().CenterPoint;
            double OffCenter, OffCenterCriteria = 80;//8cm

            if (SetupField.Any())
            {
                BeamIsocenter = SetupField.First().IsocenterPosition;
                OffCenter = BodyCenterPoint.x - BeamIsocenter.x;
                if (OffCenter > OffCenterCriteria)
                {
                    MandatorySetupField.Add("G270");
                }

                foreach (var i in MandatorySetupField)
                {
                    dummy = SetupField.Select(x => x.Id.IndexOf(i) != -1);

                    if (!dummy.Where(x => x == true).Any())
                    {
                        SetupFieldIsExist = false;
                    }
                }
            }
            else
            {
                SetupFieldIsExist = false;
            }

            return SetupFieldIsExist;
        }

        //Check Setup Field Parameters
        public bool? CheckSetupFieldParameters(PlanSetup plan)
        {
            bool? SetupFieldParameterIsCorrect = true;
            double CriteriaCollRtn = 0.0, CriteriaCouchRtn = 0.0, CriteriaY1 = -10.0, CriteriaY2 = 10.0, CriteriaX1 = -10.0, CriteriaX2 = 10.0;
            double PlanGantryAngle = 0.0, PlanCollAngle = 0.0, PlanCouchAngle = 0.0, PlanY1 = 0.0, PlanY2 = 0.0, PlanX1 = 0.0, PlanX2 = 0.0;
            double SetupFieldIdAngle = 0.0;
            IEnumerable<Beam> SetupField = plan.Beams.Where(x => x.IsSetupField == true);


            if (CheckMandatorySetupFieldIsExist(plan) == true)
            {
                foreach (var i in SetupField)
                {
                    SetupFieldIdAngle = (i.Id.IndexOf("CBCT") != -1) ? 0.0 : ExtractGantryAngleFromFieldId(i.Id, "G", 0);
                    PlanGantryAngle = i.ControlPoints.First().GantryAngle;
                    PlanCollAngle = i.ControlPoints.First().CollimatorAngle;
                    PlanCouchAngle = i.ControlPoints.First().PatientSupportAngle;
                    PlanX1 = i.ControlPoints.First().JawPositions.X1 / 10;
                    PlanX2 = i.ControlPoints.First().JawPositions.X2 / 10;
                    PlanY1 = i.ControlPoints.First().JawPositions.Y1 / 10;
                    PlanY2 = i.ControlPoints.First().JawPositions.Y2 / 10;

                    if (PlanGantryAngle != SetupFieldIdAngle || PlanCollAngle != CriteriaCollRtn || PlanCouchAngle != CriteriaCouchRtn || PlanX1 != CriteriaX1 || PlanX2 != CriteriaX2
                        || PlanY1 != CriteriaY1 || PlanY2 != CriteriaY2)
                    {
                        SetupFieldParameterIsCorrect = false;
                    }
                }
            }
            else
            {
                //SetupFieldParameterIsCorrect = null;
            }
            return SetupFieldParameterIsCorrect;
        }

        //Query Structure with Override Density
        public string QueryStructureWithDensityOverride(PlanSetup plan)
        {
            double dummy;
            IEnumerable<Structure> StructureWithDensityOverride = plan.StructureSet.Structures.Where(x => x.GetAssignedHU(out dummy) == true
            && !string.Equals(x.DicomType, "SUPPORT"));

            List<StructureWithDensity> OverrideList = new List<StructureWithDensity>();


            if (StructureWithDensityOverride.Any())
            {
                foreach (var i in StructureWithDensityOverride)
                {
                    i.GetAssignedHU(out dummy);
                    OverrideList.Add(new StructureWithDensity() { StructureId = i.Id, HUValue = dummy });
                }
            }

            string StructureOverrideList = string.Empty;
            foreach (var i in OverrideList)
            {
                StructureOverrideList += string.Format("{0}={1}\n", i.StructureId, i.HUValue);
            }
            //MessageBox.Show(StructureOverrideList);
            return StructureOverrideList;
        }

        //Check DoseRate For EBRTbeam
        public bool CheckDoseRateForEBRT(PlanSetup plan, string PlanTechnique)
        {
            IEnumerable<Beam> TreatmentField = plan.Beams.Where(x => x.IsSetupField == false);
            IEnumerable<bool> Result;
            int Criteria_Electron = 400;
            int Criteria_TBI = 200;
            int Criteria_Photon = 0;
            bool IsDoseRateCorrect = true;

            foreach (var beam in TreatmentField)
            {
                if (Equals(beam.EnergyModeDisplayName, "6X"))
                {
                    Criteria_Photon = 600;
                }
                else if (Equals(beam.EnergyModeDisplayName, "10X"))
                {
                    Criteria_Photon = 600;
                }
                else if (Equals(beam.EnergyModeDisplayName, "6X-FFF"))
                {
                    Criteria_Photon = 1400;
                }
                else if (Equals(beam.EnergyModeDisplayName, "10X-FFF"))
                {
                    Criteria_Photon = 2400;
                }
            }

            switch (PlanTechnique)
            {
                case "Electron":
                    Result = TreatmentField.Select(x => x.DoseRate == Criteria_Electron);
                    break;

                case "TBI":
                    Result = TreatmentField.Select(x => x.DoseRate == Criteria_TBI);
                    break;

                default:
                    Result = TreatmentField.Select(x => x.DoseRate == Criteria_Photon);
                    break;
            }

            foreach (var i in Result)
            {
                IsDoseRateCorrect = IsDoseRateCorrect && i;
            }

            return IsDoseRateCorrect;

        }

        //Check Whether Reference Point Has Location
        public bool CheckReferencePointNotExistLocation(PlanSetup plan)
        {
            bool IsReferencePointLocationNotExist = true;
            foreach (var i in plan.ReferencePoints)
            {
                if (i.HasLocation(plan))
                {
                    IsReferencePointLocationNotExist = false;
                    break;
                }
            }

            return IsReferencePointLocationNotExist;
        }

        //Check Multiple Reference Point
        public bool CheckWhetherMultipleReferencePointNotExist(PlanSetup plan)
        {
            bool IsMultipleReferencePointNotExist = false;
            int NumberOfReferencePoint = plan.ReferencePoints.Count();

            if (NumberOfReferencePoint == 1) IsMultipleReferencePointNotExist = true;
            return IsMultipleReferencePointNotExist;

        }

        //Check Primary Reference Point Id = PatientVolumeId        
        public bool CheckPrimaryReferencePointMatchPatientVolumeAndTargetVolumeId(PlanSetup plan)
        {
            bool IsPrimaryReferencePointMatchPatientVolumeId = false;
            string ReferencePointId = plan.PrimaryReferencePoint.Id;
            string PatientVolumeId = plan.PrimaryReferencePoint.PatientVolumeId;
            string TargetVolumeId = plan.TargetVolumeID;

            if (string.Equals(ReferencePointId, PatientVolumeId) && string.Equals(ReferencePointId, TargetVolumeId)) { IsPrimaryReferencePointMatchPatientVolumeId = true; }
            return IsPrimaryReferencePointMatchPatientVolumeId;
        }

        //Check Primary Reference Point DoseLimit
        public bool CheckPrimaryReferencePointDoseLimit(PlanSetup plan)
        {
            bool IsPrimaryReferencePointDoseLimitCorrect = false;
            double TotalDose = (plan.TotalDose.IsAbsoluteDoseValue) ? Math.Round(plan.TotalDose.Dose) : double.NaN;
            double DosePerFraction = (plan.DosePerFraction.IsAbsoluteDoseValue) ? Math.Round(plan.DosePerFraction.Dose) : double.NaN;
            double Fraction = (DosePerFraction == double.NaN) ? Convert.ToInt32(TotalDose / DosePerFraction) : double.NaN;

            double TotalDoseLimit = Math.Round(plan.PrimaryReferencePoint.TotalDoseLimit.Dose);
            double DailyDoseLimit = Math.Round(plan.PrimaryReferencePoint.DailyDoseLimit.Dose);
            double SessionDoseLimit = Math.Round(plan.PrimaryReferencePoint.SessionDoseLimit.Dose);

            if (TotalDose == TotalDoseLimit && DosePerFraction == DailyDoseLimit && DosePerFraction == SessionDoseLimit)
            {
                IsPrimaryReferencePointDoseLimitCorrect = true;
            }
            else if (plan.ApprovalStatus == PlanSetupApprovalStatus.TreatmentApproved)
            {
                string message = string.Format("Total Dose Limit = {0}\nSession Dose Limit = {1}\nDaily Dose Limit= {2}", TotalDoseLimit, SessionDoseLimit, DailyDoseLimit);
                string caption = "Please Check Dose Limit";
                MessageBoxButton buttons = MessageBoxButton.YesNo;
                MessageBoxResult result;

                result = MessageBox.Show(message, caption, buttons);

                if (result == MessageBoxResult.Yes) { IsPrimaryReferencePointDoseLimitCorrect = true; }
            }
            return IsPrimaryReferencePointDoseLimitCorrect;

        }

        //Check Tolerence Table
        public bool CheckTolerenceTableForEBRT(PlanSetup plan, string plantechnique)
        {
            bool IsToleranceTableCorrect = false;
            string Criteria = string.Empty;
            int SetupFieldNum = 0, TreatmentBeamNumber = 0, BeamNumber = plan.Beams.Count();
            int CorrectBeamAccumulator = 0;

            if (!String.Equals(plantechnique, "Electron"))
            {
                Criteria = "Photon";
            }
            else
            {
                Criteria = "Electron";
            }

            foreach (var eachbeam in plan.Beams)
            {
                if (!eachbeam.IsSetupField)
                {
                    if (String.Equals(eachbeam.ToleranceTableLabel, Criteria))
                    {
                        CorrectBeamAccumulator++;
                    }
                }
                else { SetupFieldNum++; }
            }
            TreatmentBeamNumber = BeamNumber - SetupFieldNum;

            if (CorrectBeamAccumulator == TreatmentBeamNumber)
            {
                IsToleranceTableCorrect = true;
            }
            else { IsToleranceTableCorrect = false; }

            return IsToleranceTableCorrect;
        }

        //Check Structure Type as Target Volume
        public bool CheckTargetVolumeType(PlanSetup plan)
        {
            string TargetVolumeName = plan.TargetVolumeID;
            string[] PassingDicomType = new string[] { "PTV", "CTV", "GTV" };
            bool TargetVoumeTypeIsCorrect = false;
            //string msg = String.Empty;
            foreach (var structure in plan.StructureSet.Structures)
            {
                if (String.Equals(structure.Id, TargetVolumeName))
                {
                    foreach (var DicomType in PassingDicomType)
                    {
                        if (String.Equals(structure.DicomType, DicomType)) TargetVoumeTypeIsCorrect = true;
                    }
                    break;
                }
            }
            //MessageBox.Show(msg);

            return TargetVoumeTypeIsCorrect;
        }

        //Check Beam Isocenter: 比較Beam Center與最近Marker位置
        public bool? CheckBeamCenter(PlanSetup plan)
        {
            VVector BeamCenterPosition = new VVector(0.0, 0.0, 0.0);
            VVector UserOriginPosition = plan.StructureSet.Image.UserOrigin;
            VVector UserOriginToBeamCenter = new VVector();

            //String msg = String.Empty;
            double Criteria = 0.04;
            bool? IsBeamCenterOnCorrectPosition = null;

            foreach (var eachbeam in plan.Beams)
            {
                if (!eachbeam.IsSetupField)
                {
                    BeamCenterPosition = eachbeam.IsocenterPosition;
                    UserOriginToBeamCenter = BeamCenterPosition - UserOriginPosition;
                    if (Math.Abs(UserOriginToBeamCenter.x) <= Criteria && Math.Abs(UserOriginToBeamCenter.y) <= Criteria
                        && Math.Abs(UserOriginToBeamCenter.z) <= Criteria)
                    {
                        IsBeamCenterOnCorrectPosition = true;
                    }
                    else
                    {
                        VVector ProximalMarkerToBeamCenter = new VVector();
                        VVector ProximalMarker = FindProximalMarkerBetweenInputCenter(plan, BeamCenterPosition);

                        if (Double.IsNaN(ProximalMarker.x) != true && Double.IsNaN(ProximalMarker.y) != true && Double.IsNaN(ProximalMarker.z) != true)
                        {
                            ProximalMarkerToBeamCenter = BeamCenterPosition - ProximalMarker;
                            if (ProximalMarkerToBeamCenter.x <= Criteria && ProximalMarkerToBeamCenter.y <= Criteria && ProximalMarkerToBeamCenter.z <= Criteria)
                            {
                                IsBeamCenterOnCorrectPosition = true;
                            }
                            else
                            {
                                IsBeamCenterOnCorrectPosition = false;
                            }
                        }
                        else
                        { IsBeamCenterOnCorrectPosition = null; }
                        //msg += String.Format("{0},{1},{2}\n", Double.IsNaN(ProximalMarker.x), Double.IsNaN(ProximalMarker.y), Double.IsNaN(ProximalMarker.z));
                    }

                }
            }
            //MessageBox.Show(msg);
            return IsBeamCenterOnCorrectPosition;

        }

        public VVector FindProximalMarkerBetweenInputCenter(PlanSetup plan, VVector center)
        {
            String DicomType = "MARKER";
            double Distance, ProximalDistance = 1000000000;
            int markercounter = 0;
            VVector ProximalMarkerPosition = new VVector();

            foreach (var structure in plan.StructureSet.Structures)
            {
                if (Equals(structure.DicomType, DicomType))
                {
                    Distance = VVector.Distance(center, structure.CenterPoint);
                    if (ProximalDistance >= Distance)
                    {
                        ProximalDistance = Distance;
                        ProximalMarkerPosition = structure.CenterPoint;
                    }

                    markercounter++;
                }
            }

            if (markercounter == 0) ProximalMarkerPosition = new VVector(Double.NaN, Double.NaN, Double.NaN);
            return ProximalMarkerPosition;
        }

        //Check User Origin: 比較User Origin與Marker的位置，與確認User Origin in-out在影像中心。
        public bool? CheckUserOrigin(PlanSetup plan)
        {
            // The Unit of Structure.CenterPoint and UserOrigin is in mm;
            //String msg = String.Empty;
            double DistanceCriteria = Math.Pow(10, 0);//鉛線橫切面的寬度約2mm
            double PositionCriteria_InOut = 0.0;
            double ProximalDistance = 1000000000;
            VVector UserOriginPosition = new VVector();
            VVector ProximalMarkerPosition = new VVector(0.0, 0.0, 0.0);
            bool? IsUserOriginOnCorrectPosition = null;

            UserOriginPosition = plan.StructureSet.Image.UserOrigin;
            ProximalMarkerPosition = FindProximalMarkerBetweenInputCenter(plan, UserOriginPosition);
            ProximalDistance = VVector.Distance(UserOriginPosition, ProximalMarkerPosition);

            //msg += String.Format("{0}", ProximalDistance);
            //MessageBox.Show(msg);

            if (UserOriginPosition.z == PositionCriteria_InOut)
            {
                if (ProximalDistance <= DistanceCriteria)
                {
                    IsUserOriginOnCorrectPosition = true;
                }
                else if (ProximalDistance > DistanceCriteria && ProximalDistance != 1000000000)
                {
                    IsUserOriginOnCorrectPosition = false;
                }
                else if (ProximalDistance == 1000000000)
                {
                    IsUserOriginOnCorrectPosition = null;
                }
            }
            else
            {
                IsUserOriginOnCorrectPosition = false;
            }

            return IsUserOriginOnCorrectPosition;
        }

        //Check是否存在HyperArcCouch
        public bool CheckHyperArcCouch(PlanSetup plan)
        {
            string couchtop = "Encompass", couchbase = "Encompass_Base";

            bool Encompass = false, EncompassBase = false;
            double EncompassHU = 400, EncompassBaseHU = -400, structureHU;

            foreach (var structure in plan.StructureSet.Structures)
            {
                structure.GetAssignedHU(out structureHU);
                if (Equals(structure.DicomType, "SUPPORT"))
                {
                    foreach (var i in structure.StructureCodeInfos)
                    {
                        if (structureHU == EncompassHU && Equals(i.Code, couchtop)) Encompass = true;
                        if (structureHU == EncompassBaseHU && Equals(i.Code, couchbase)) EncompassBase = true;
                    }
                }
            }

            if (Encompass == true && EncompassBase == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        //Check是否存在IGRTMediumCouch
        public bool CheckIGRTCouch(PlanSetup plan)
        {
            string couchtype = "Exact IGRT Couch, medium";
            bool CouchInterior = false, CouchSurface = false;
            double CouchInteriorHU = -950, CouchSurfaceHU = -550, structureHU;

            foreach (var structure in plan.StructureSet.Structures)
            {
                structure.GetAssignedHU(out structureHU);
                if (Equals(structure.DicomType, "SUPPORT") && Equals(structure.Name, couchtype))
                {
                    if (structureHU == CouchInteriorHU) CouchInterior = true;
                    if (structureHU == CouchSurfaceHU) CouchSurface = true;
                }
            }

            if (CouchInterior == true && CouchSurface == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        //建立病人相關基本資料
        public GeneralInformation[] GetGeneralInformation(ScriptContext context, List<PlanSetup> plan, List<string> plantype, List<string> plantechnique)
        {

            GeneralInformation[] generalinformation = new GeneralInformation[plan.Count()];

            for (var i = 0; i < plan.Count(); i++)//loop 每一個plan
            {
                GeneralInformation tmparray = new GeneralInformation();
                tmparray.PatientName = context.Patient.Name;
                tmparray.PlanName = plan[i].Id;
                tmparray.Course = plan[i].Course.Id;
                tmparray.ImageName = (context.Image == null) ? plan[i].StructureSet.Image.Id : context.Image.Id;
                tmparray.IsHyperArc = IsHyperArc(plan[i]);
                tmparray.StructureSetName = (context.StructureSet == null) ? plan[i].StructureSet.Id : context.StructureSet.Id;
                tmparray.TotalDose = plan[i].TotalDose.Dose;
                tmparray.Fraction = plan[i].TotalDose.Dose / plan[i].DosePerFraction.Dose;
                tmparray.Scanner = (context.Image == null) ? plan[i].StructureSet.Image.Series.ImagingDeviceId : context.Image.Series.ImagingDeviceId;
                tmparray.PlanType = plantype[i];
                tmparray.PlanTechnique = plantechnique[i];
                tmparray.StructureOverride = QueryStructureWithDensityOverride(plan[i]);
                tmparray.PlanUid = plan[i].UID;

                IEnumerable<Beam> TreatmentBeam = plan[i].Beams.Where(x => x.IsSetupField == false);
                tmparray.TreatmentUnit = (CheckTreatmentUnitConsistency(plan[i])) ? TreatmentBeam.First().TreatmentUnit.Id : string.Empty;

                tmparray.ApprovalStatus = Convert.ToString(plan[i].ApprovalStatus);

                generalinformation[i] = tmparray;
            }
            return generalinformation;
        }

        //check是否是HyperArc
        public bool IsHyperArc(PlanSetup plan)
        {
            bool ishyperarc = false;
            int SetupNum = 0, HyperArcBeamNum = 0;
            foreach (var i in plan.Beams)
            {
                if (!i.IsSetupField)
                {
                    if (Equals(Convert.ToString(i.SetupTechnique), "Unknown"))
                    {
                        HyperArcBeamNum++;
                    }
                }
                else
                {
                    SetupNum++;
                }

            }

            if (HyperArcBeamNum == plan.Beams.Count() - SetupNum)
            {
                ishyperarc = true;
            }
            return ishyperarc;
        }

        //check Plan是使用VMAT, IMRT, FiF, 3or2D, 3or2DwithoutMLC, Electron, TBI
        public List<string> EBRTCheckTechnique(List<PlanSetup> plan)
        {
            //string msg = String.Empty;
            string technique = string.Empty;
            //int accumulator = 0;
            List<string> TechniqueResult = new List<string>();

            bool IsVMAT = true;
            bool IsElectron = true;
            bool IsFiF = true;
            bool Is3or2D = true;
            bool IS3or2DwithoutMLC = true;
            bool IsIMRT = true;
            bool IsTBI = true;

            //string msg = string.Empty;
            foreach (var eachplan in plan)
            {
                int BeamNumber = eachplan.Beams.Count(), SetupNumber = 0, TreatmentBeamNumber;
                //[VMAT,Electron,FiF,3or2D,3or2DwithoutMLC,IMRT,TBI] number of current technique
                int[] technique_accumulator = new int[7] { 0, 0, 0, 0, 0, 0, 0 };

                foreach (var eachbeam in eachplan.Beams)
                {

                    if (!eachbeam.IsSetupField)
                    {
                        IsVMAT = true;
                        IsElectron = true;
                        IsFiF = true;
                        Is3or2D = true;
                        IS3or2DwithoutMLC = true;
                        IsIMRT = true;
                        IsTBI = true;


                        //msg += string.Format("{0},{1}\n", Equals(eachbeam.Technique.Id, "SRS ARC"), Equals(eachbeam.Technique.Id, "ARC"));

                        if (Equals(eachbeam.Technique.Id, "ARC") || Equals(eachbeam.Technique.Id, "SRS ARC"))
                        {
                            IsElectron = false;
                            IsFiF = false;
                            Is3or2D = false;
                            IS3or2DwithoutMLC = false;
                            IsIMRT = false;
                            IsTBI = false;
                        }

                        if (Equals(eachbeam.Technique.Id, "STATIC") || Equals(eachbeam.Technique.Id, "SRS STATIC"))
                        {
                            IsTBI = false;
                            IsVMAT = false;

                            //利用FiF每個subfield的照野重疊的關係
                            //i的範圍各縮1，因FiF第一個與最後一個control point不重複。
                            //i=i+2 每兩個control point比較一次weighting
                            if (eachbeam.ControlPoints.Count() > 2)
                            {
                                Is3or2D = false;
                                IS3or2DwithoutMLC = false;
                                IsElectron = false;
                                for (var i = 0 + 1; i < eachbeam.ControlPoints.Count() - 1; i = i + 2)
                                {
                                    if (!Equals(eachbeam.ControlPoints[i].MetersetWeight, eachbeam.ControlPoints[i + 1].MetersetWeight))
                                    {
                                        IsFiF = false;
                                    }
                                    else
                                    {
                                        IsIMRT = false;
                                    }

                                }
                            }
                            else
                            {
                                IsFiF = false;
                                IsIMRT = false;

                                if (eachbeam.EnergyModeDisplayName.IndexOf("E") == -1)
                                {
                                    IsElectron = false;

                                    if (!Equals(Convert.ToString(eachbeam.MLCPlanType), "NotDefined"))
                                    {
                                        IS3or2DwithoutMLC = false;
                                    }
                                    else
                                    {
                                        Is3or2D = false;
                                    }

                                }
                                else
                                {
                                    Is3or2D = false;
                                    IS3or2DwithoutMLC = false;
                                }
                            }
                        }

                        if (Equals(eachbeam.Technique.Id, "TOTAL"))
                        {
                            IsVMAT = false;
                            IsElectron = false;
                            IsFiF = false;
                            Is3or2D = false;
                            IS3or2DwithoutMLC = false;
                            IsIMRT = false;
                        }

                        //判斷每一隻beam的技術
                        if (IsVMAT)
                        {
                            technique_accumulator[0] = technique_accumulator[0] + 1;
                        }
                        if (IsElectron)
                        {
                            technique_accumulator[1] = technique_accumulator[1] + 1;
                        }
                        if (IsFiF)
                        {
                            technique_accumulator[2] = technique_accumulator[2] + 1;
                        }
                        if (Is3or2D)
                        {
                            technique_accumulator[3] = technique_accumulator[3] + 1;
                        }
                        if (IS3or2DwithoutMLC)
                        {
                            technique_accumulator[4] = technique_accumulator[4] + 1;
                        }
                        if (IsIMRT)
                        {
                            technique_accumulator[5] = technique_accumulator[5] + 1;
                        }
                        if (IsTBI)
                        {
                            technique_accumulator[6] = technique_accumulator[6] + 1;
                        }
                    }
                    else
                    {
                        SetupNumber++;
                    }
                }

                TreatmentBeamNumber = BeamNumber - SetupNumber;

                //確認plan內所有射束是否都是同一種技術才確認該plan為何種技術。例外為FiF與3Dmix的plan會顯示FiF。其餘會回傳空白。
                if (technique_accumulator[0] == TreatmentBeamNumber)
                {
                    TechniqueResult.Add("VMAT");
                }
                if (technique_accumulator[1] == TreatmentBeamNumber)
                {
                    TechniqueResult.Add("Electron");
                }
                if (technique_accumulator[2] == TreatmentBeamNumber)
                {
                    TechniqueResult.Add("FiF");
                }
                if (technique_accumulator[3] == TreatmentBeamNumber)
                {
                    TechniqueResult.Add("3or2D");
                }
                if (technique_accumulator[4] == TreatmentBeamNumber)
                {
                    TechniqueResult.Add("3or2DwithoutMLC");
                }
                if (technique_accumulator[5] == TreatmentBeamNumber)
                {
                    TechniqueResult.Add("IMRT");
                }
                if (technique_accumulator[6] == TreatmentBeamNumber)
                {
                    TechniqueResult.Add("TBI");
                }
                if (technique_accumulator[2] != 0 && technique_accumulator[3] != 0)
                {
                    TechniqueResult.Add("FiF");
                }

            }

            //msg += string.Format("{0}\n", TechniqueResult.Count);
            if (TechniqueResult.Count > 1 || TechniqueResult.Count == 0)
            {
                TechniqueResult.Clear();
                TechniqueResult.Add(null);
                MessageBox.Show("Some wrongs happened in EBRTCheckTechnique function", "Error");
            }
            //MessageBox.Show(msg);
            return TechniqueResult;

        }

        //check PlanSetup是EBRT, Brachy or EBRTProton
        public List<string> CheckPlanType(List<PlanSetup> plan)
        {
            List<string> Result = new List<string>();
            foreach (var eachplan in plan)
            {
                if (Equals(Convert.ToString(eachplan.PlanType), "ExternalBeam"))
                {
                    Result.Add("ExternalBeam");
                }
                else if (Equals(Convert.ToString(eachplan.PlanType), "Brachy"))
                {
                    Result.Add("Brachy");
                }
                else if (Equals(Convert.ToString(eachplan.PlanType), "ExternalBeam_Proton"))
                {
                    Result.Add("ExternalBeam_Proton");
                }
            }
            return Result;

        }

        //check 開啟的Plan是EBRTplan還是PlanSum，如果超過一個PlanSum則會要求選擇。
        //執行後plan(PlanSetup)會是目前開啟的plan或是plansum的PlanSetup
        public static Window PlanSumSelect { get; set; }
        public List<PlanSetup> CurrentPlan(List<PlanSetup> plan)
        {
            if (context.PlanSetup != null)
            {
                plan.Add(context.PlanSetup);
                //MessageBox.Show("One Plan");
            }
            else if (context.Course == null)
            {
                MessageBox.Show("No Course");
            }
            else if (context.Course.PlanSums.Count() != 0)
            {
                //MessageBox.Show(Convert.ToString(context.Course.PlanSums.Count()));
                string msg = string.Empty;
                int NumberOfSum = context.Course.PlanSums.Count();
                List<PlanSum> plansum = new List<PlanSum>();
                plansum = context.Course.PlanSums.ToList();

                if (NumberOfSum > 1)// more than one plansum in the course
                {
                    PlanSumSelect = new Window() { Height = 200, Width = 400 };
                    StackPanel stack1 = new StackPanel() { Orientation = Orientation.Vertical };
                    Label label = new Label() { Content = "Select Plansum", FontSize = 36, HorizontalAlignment = HorizontalAlignment.Center };
                    ComboBox PlanSumSelectCombo = new ComboBox()
                    {
                        ItemsSource = plansum,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Height = 70,
                        Width = 200,
                        FontSize = 36
                    };
                    PlanSumSelectCombo.DropDownClosed += new EventHandler(ClosePlanSumSelectWindow);
                    stack1.Children.Add(label);
                    stack1.Children.Add(PlanSumSelectCombo);
                    PlanSumSelect.Content = stack1;

                    PlanSumSelect.ShowDialog();

                    foreach (var i in plansum)
                    {
                        if (Equals(i.Id, Convert.ToString(PlanSumSelectCombo.SelectedItem)))
                        {
                            plan = i.PlanSetups.ToList();
                            //MessageBox.Show(plan[1].Id);
                        }
                    }
                }
                else
                {
                    foreach (var i in plansum)
                        plan = i.PlanSetups.ToList();
                }
            }
            else
            {
                MessageBox.Show("No Plan is loaded or no plansum in the course");
            }

            return plan;
        }

        //關閉CurrentPlan內Window: PlanSumSelect視窗
        public void ClosePlanSumSelectWindow(object sender, EventArgs e)
        {
            PlanSumSelect.Close();
        }


    }
}
