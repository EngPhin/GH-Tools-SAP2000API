using System;
using System.Collections;
using System.Collections.Generic;

using Rhino;
using Rhino.Geometry;

using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;


using SAP2000v1;
using System.Linq;




/// <summary>
/// This class will be instantiated on demand by the Script component.
/// </summary>
public class Script_Instance : GH_ScriptInstance
{
#region Utility functions
  /// <summary>Print a String to the [Out] Parameter of the Script component.</summary>
  /// <param name="text">String to print.</param>
  private void Print(string text) { /* Implementation hidden. */ }
  /// <summary>Print a formatted String to the [Out] Parameter of the Script component.</summary>
  /// <param name="format">String format.</param>
  /// <param name="args">Formatting parameters.</param>
  private void Print(string format, params object[] args) { /* Implementation hidden. */ }
  /// <summary>Print useful information about an object instance to the [Out] Parameter of the Script component. </summary>
  /// <param name="obj">Object instance to parse.</param>
  private void Reflect(object obj) { /* Implementation hidden. */ }
  /// <summary>Print the signatures of all the overloads of a specific method to the [Out] Parameter of the Script component. </summary>
  /// <param name="obj">Object instance to parse.</param>
  private void Reflect(object obj, string method_name) { /* Implementation hidden. */ }
#endregion

#region Members
  /// <summary>Gets the current Rhino document.</summary>
  private readonly RhinoDoc RhinoDocument;
  /// <summary>Gets the Grasshopper document that owns this script.</summary>
  private readonly GH_Document GrasshopperDocument;
  /// <summary>Gets the Grasshopper script component that owns this script.</summary>
  private readonly IGH_Component Component;
  /// <summary>
  /// Gets the current iteration count. The first call to RunScript() is associated with Iteration==0.
  /// Any subsequent call within the same solution will increment the Iteration count.
  /// </summary>
  private readonly int Iteration;
#endregion

  /// <summary>
  /// This procedure contains the user code. Input parameters are provided as regular arguments,
  /// Output parameters as ref arguments. You don't have to assign output parameters,
  /// they will have a default value.
  /// </summary>
  private void RunScript(System.Object cSAP, bool Active, List<string> DefineLoadCombo, List<string> nodeNumbers, ref object ActiveNode, ref object LoadCombos, ref object LC_F1, ref object LC_F2, ref object LC_F3)
  {

    // use ret to test results
    // Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero)


    if(!Active)
    {
      Component.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Component is not set to active");
      return;
    }
    //Castmodel
    var mySapModel = CastCsapModel(cSAP);

    //Enquire all node
    var nodeNo = mySapModel.PointElm.Count();

    //Enquire number of frames
    var allFrameCount = mySapModel.FrameObj.Count();

    double[] SapResult = new double[7];
    // ret = mySapModel.FrameObj.GetPoints(FrameName[1], ref temp_string1, ref temp_string2);
    // PointName[0] = temp_string1;
    // PointName[1] = temp_string2;

    //get SAP2000 results
    int NumberResults = 1;
    string[] Obj = new string[1];
    string[] Elm = new string[1];
    string[] LoadCombo = new string[216];
    string[] StepType = new string[2];
    double[] StepNum = new double[1];
    double[] U1 = new double[1];
    double[] U2 = new double[1];
    double[] U3 = new double[1];
    double[] R1 = new double[1];
    double[] R2 = new double[1];
    double[] R3 = new double[1];
    double[] F1 = new double[1];
    double[] F2 = new double[1];
    double[] F3 = new double[1];
    double[] M1 = new double[1];
    double[] M2 = new double[1];
    double[] M3 = new double[1];

    List<double> jointDispListU1 = new List<double>();
    List<double> jointDispListU2 = new List<double>();
    List<double> jointReactListF1 = new List<double>();
    List<double> jointReactListF2 = new List<double>();
    List<double> jointReactListF3 = new List<double>();
    List<string> nodeNumberList = new List<string>();
    List<string> loadComboList = new List<string>();

    int ret = 1;
    int ret1 = 1;
    int ret2 = 1;
    int n = (DefineLoadCombo.Count());

    int numberOfNodes = nodeNumbers.Count();
    int numberofLoadCombination = mySapModel.RespCombo.Count();

    for (var j = 0;j < numberOfNodes;j++)
    {
      string jString = nodeNumbers[j];
      string nodeNumber = jString;

      for (int k = 0; k < n; k++)
      {
        string loadCombo = DefineLoadCombo[k];

        ret1 = mySapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
        ret1 = mySapModel.Results.Setup.SetComboSelectedForOutput(loadCombo);
        ret = mySapModel.Results.JointDispl(nodeNumber, eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCombo, ref StepType, ref StepNum, ref U1, ref U2, ref U3, ref R1, ref R2, ref R3);
        ret2 = mySapModel.Results.JointReact(nodeNumber, eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCombo, ref StepType, ref StepNum, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3);
        jointDispListU1.Add(U1[0]);
        jointDispListU2.Add(U2[0]);
        jointReactListF1.Add(F1[0]);
        jointReactListF2.Add(F2[0]);
        jointReactListF3.Add(F3[0]);
        nodeNumberList.Add(nodeNumber);
        loadComboList.Add(loadCombo);
      }

    }

    //    LC_U1 = jointDispListU1;
    //    LC_U2 = jointDispListU2;
    //LoadCaseCount = numberofLoadCombination.ToString() ;
    LoadCombos = loadComboList;
    ActiveNode = nodeNumberList;
    LC_F1 = jointReactListF1;
    LC_F2 = jointReactListF2;
    LC_F3 = jointReactListF3;
    //Ret=ret;
    //Ret1=ret1;
    //Ret2=ret2;
    //    A = nodeNo;



  }

  // <Custom additional code> 


  private static cSapModel CastCsapModel(object cSap)
  {
    cSapModel mySapModel;
    try
    {
      mySapModel = (SAP2000v1.cSapModel) cSap;
    }
    catch (InvalidCastException)
    {
      var msg = "Unable to cast cSap";
      throw new Exception(msg);
    }
    return mySapModel;
  }

  // </Custom additional code> 
}