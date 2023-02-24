using System;
using System.Collections;
using System.Collections.Generic;

using Rhino;
using Rhino.Geometry;

using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;


using System.IO;
using System.Linq;
using System.Windows.Forms;
using SAP2000v1;
using System.Text;


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
  private void RunScript(string FilePath, bool Active, ref object cSAP)
  {

    if (!Active)
    {
      Component.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Component is not set to active");
      return;
    }

    //set the following flag to true to attach to an existing instance of the program
    //otherwise a new instance of the program will be started
    bool AttachToInstance;
    AttachToInstance = false;

    //set the following flag to true to manually specify the path to SAP2000.exe
    //this allows for a connection to a version of SAP2000 other than the latest installation
    //otherwise the latest installed version of SAP2000 will be launched
    bool SpecifyPath;
    SpecifyPath = false;

    //if the above flag is set to true, specify the path to SAP2000 below
    string ProgramPath;
    ProgramPath = "C:\\Program Files (x86)\\Computers and Structures\\SAP2000 20\\SAP2000.exe";

    //full path to the model
    //set it to the desired path of your model
    string ModelDirectory = FilePath;
    //var DirectoryExpanded = CheckPath(GetModelDirectory.FilePath);

    //The model has existed. These lines are not needed
    // try
    // {
    //     System.IO.Directory.CreateDirectory(ModelDirectory);
    // }
    // catch (Exception ex)
    // {
    //     Console.WriteLine("Could not create directory: " + ModelDirectory);
    // }

    //string ModelName = ".Testsdb";
    string ModelPath = ModelDirectory; //+ System.IO.Path.DirectorySeparatorChar + ModelName;

    //dimension the SapObject as cOAPI type
    cOAPI mySapObject = null;

    //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero)
    int ret = 0;

    if (AttachToInstance)
    {
      //attach to a running instance of SAP2000
      try
      {
        //get the active SapObject
        mySapObject = (cOAPI) System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapObject");
      }
      catch (Exception ex)
      {
        Console.WriteLine("No running instance of the program found or failed to attach.");
        return;
      }
    }
    else
    {
      //create API helper object
      cHelper myHelper;
      try
      {
        myHelper = new Helper();
      }
      catch (Exception ex)
      {
        Console.WriteLine("Cannot create an instance of the Helper object");
        return;
      }
      if (SpecifyPath)
      {
        //'create an instance of the SapObject from the specified path
        try
        {
          //create SapObject
          mySapObject = myHelper.CreateObject(ProgramPath);
        }
        catch (Exception ex)
        {
          Console.WriteLine("Cannot start a new instance of the program from " + ProgramPath);
          return;
        }
      }
      else
      {
        //'create an instance of the SapObject from the latest installed SAP2000
        try
        {
          //create SapObject
          mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject");
        }
        catch (Exception ex)
        {
          Console.WriteLine("Cannot start a new instance of the program.");
          return;
        }
      }
      //start SAP2000 application
      ret = mySapObject.ApplicationStart();
    }

    //create SapModel object
    cSapModel mySapModel;
    mySapModel = mySapObject.SapModel;

    ret = mySapModel.InitializeNewModel();
    ret = mySapModel.File.OpenFile(ModelPath);
    cSAP = mySapModel;
  }

  // <Custom additional code> 



  // </Custom additional code> 
}