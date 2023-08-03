Attribute VB_Name = "TestRoutines"
'Finite Element Method, Problem 3.27
'This example tests the program's ability to correctly model 2D truss elements, sum forces at joints,
'report axial forces, and report deflections at nodes
'The test was successfully run on 8/16/2017
Private Sub FEM_Problem_3_27()
    
    'Start a new finite element model
    Dim myModel As New FEModel
    
    'Define the nodes
    Call myModel.AddNode("N1", 0, 0)
    Call myModel.AddNode("N2", 60 * 12, 0)
    Call myModel.AddNode("N3", 30 * 12, 40 * 12)
    Call myModel.AddNode("N4", 30 * 12, 60 * 12)
    
    'Define the members
    Call myModel.AddMember("M1", "N1", "N3", 30000000, 10, 3)
    Call myModel.AddMember("M2", "N2", "N3", 30000000, 10, 3)
    Call myModel.AddMember("M3", "N3", "N4", 30000000, 10, 3)
    
    'Define the supports
    Call myModel.EditSupport("N1", True, True, True)
    Call myModel.EditSupport("N2", True, True, True)
    Call myModel.EditSupport("N4", True, True, True)
    
    'Define the member end releases
    Call myModel.EditEndReleases("M1", True, True)
    Call myModel.EditEndReleases("M2", True, True)
    Call myModel.EditEndReleases("M3", False, True)
    
    'Define the joint loads
    Call myModel.AddNodeLoad("N3", 5000, FX, "D")
    Call myModel.AddNodeLoad("N3", -10000, FY, "D")
    
    'Define a load combination (unfactored dead load in this case)
    Call myModel.AddLoadCombo("Combo 1", Array("D"), Array(1))
    
    'Clear 'Sheet1'
    Sheet1.Cells.Clear
    
    'Print the displacements at 'N3'
    Sheet1.Range("A1") = "Node N3 Displacements"
    Sheet1.Range("A2") = myModel.GetDisp("N3", DX, "Combo 1")
    Sheet1.Range("A3") = myModel.GetDisp("N3", DY, "Combo 1")
    Sheet1.Range("A4") = myModel.GetDisp("N3", RZ, "Combo 1")
    
    'Show equilibrium at 'N3' - Reactions should equal 0 at 'N3')
    Sheet1.Range("A6") = "Sum of Forces and Moments at N3"
    Sheet1.Range("A7") = myModel.GetReaction("N3", FX)
    Sheet1.Range("A8") = myModel.GetReaction("N3", FY)
    Sheet1.Range("A9") = myModel.GetReaction("N3", MZ)
    
    'Show the member axial forces
    Sheet1.Range("E1") = "Member Axial Forces"
    Sheet1.Range("E2") = myModel.GetMaxAxial("M1")
    Sheet1.Range("F2") = myModel.GetMinAxial("M1")
    Sheet1.Range("E3") = myModel.GetMaxAxial("M2")
    Sheet1.Range("F3") = myModel.GetMinAxial("M2")
    Sheet1.Range("E4") = myModel.GetMaxAxial("M3")
    Sheet1.Range("F4") = myModel.GetMinAxial("M3")
    
End Sub

'Finite Element Method, Example 4.10, Pages 196-198
'This example tests the program's ability to correctly apply end releases to members
'This test was successfully run on 8/16/2017
Private Sub FEM_Example_4_10()
    
    'Start a new finite element model
    Dim myModel As New FEModel
    
    'Define the nodes
    Call myModel.AddNode("N1", 0, 0)
    Call myModel.AddNode("N2", 6 * 12, 0)
    Call myModel.AddNode("N3", 10 * 12, 0)
    
    'Define the members
    Call myModel.AddMember("M1", "N1", "N2", 29000, 100, 10)
    Call myModel.AddMember("M2", "N2", "N3", 29000, 100, 10)
    
    'Define the supports
    Call myModel.EditSupport("N1", True, True, True)
    Call myModel.EditSupport("N3", True, True, True)
    
    'Define the member end releases
    Call myModel.EditEndReleases("M1", False, True)
    
    'Define the joint load
    Call myModel.AddNodeLoad("N2", -5, FY, "L")
    
    'Define load combinations
    Call myModel.AddLoadCombo("Live", Array("L"), Array("1"))
    
    'Clear 'Sheet1'
    Sheet1.Cells.Clear
    
    'Get the stiffness matrices for 'M1' and 'M2'
    Dim Stiffness1 As Matrix, Stiffness2 As Matrix
    Set Stiffness1 = myModel.Members("M1").LocalStiff
    Set Stiffness2 = myModel.Members("M2").LocalStiff
    
    'Remove the axial load degrees of freedom from the matrix
    Call Stiffness1.RemoveRow(4)
    Call Stiffness1.RemoveRow(1)
    Call Stiffness1.RemoveCol(4)
    Call Stiffness1.RemoveCol(1)
    
    Call Stiffness2.RemoveRow(4)
    Call Stiffness2.RemoveRow(1)
    Call Stiffness2.RemoveCol(4)
    Call Stiffness2.RemoveCol(1)
    
    'Print the stiffness matrix for 'M1'
    Sheet1.Range("A1") = "Member M1 Local Stiffness Matrix"
    Call Stiffness1.PrintMatrix(Sheet1.Range("A2"))
    
    'Print the stiffness matrix for 'M2'
    Sheet1.Range("F1") = "Member M2 Local Stiffness Matrix"
    Call Stiffness2.PrintMatrix(Sheet1.Range("F2"))
    
    'Print the displacements at "N2"
    Sheet1.Range("A7") = "Node N2 Displacements"
    Sheet1.Range("A8") = myModel.GetDisp("N2", DY)
    Sheet1.Range("A9") = myModel.GetDisp("N2", RZ)
    
    'Get the local end forces for "M1"
    Dim EndForces As Matrix
    Set EndForces = myModel.Members("M1").LocalForces("Live")
    
    'Remove the axial load degrees of freedom from the vector
    Call EndForces.RemoveRow(4)
    Call EndForces.RemoveRow(1)
    
    'Print the local end forces for 'M1'
    Sheet1.Range("A11") = "Member M1 Local End Forces"
    Call EndForces.PrintMatrix(Sheet1.Range("A12"))
    
End Sub

'Structural Analysis, Example 5.15, Pages 212-217
'This example tests the following:
    '- Axial distributed loads
    '- Transverse distributed loads
    '- End releases
    '- Member maximum and minimum moment calculations
    '- Member maximum and minimum shear calculations
    '- Member moment diagrams
    '- Member shear diagrams
    '- Member axial force diagrams
'This test was successfully run on 10/21/2017
Private Sub SA_Example_5_15()
    
    'Start a new finite element model
    Dim myModel As New FEModel
    
    'Define the nodes
    Call myModel.AddNode("A", 0, 0)
    Call myModel.AddNode("B", 0, 5)
    Call myModel.AddNode("C", 4, 8)
    Call myModel.AddNode("D", 8, 5)
    Call myModel.AddNode("E", 8, 0)
    
    'Define the members
    Call myModel.AddMember("M1", "A", "B", 1000, 500, 30)
    Call myModel.AddMember("M2", "B", "C", 1000, 500, 30)
    Call myModel.AddMember("M3", "C", "D", 1000, 500, 30)
    Call myModel.AddMember("M4", "E", "D", 1000, 500, 30)
    
    'Define the end releases
    Call myModel.EditEndReleases("M2", False, True)
    
    'Define the supports
    Call myModel.EditSupport("A", True, True, False)
    Call myModel.EditSupport("E", True, True, False)
    
    'Add member distributed loads
    Call myModel.AddMemberDistLoad("M2", 7.68, 7.68, , , Transverse, "Case 1")
    Call myModel.AddMemberDistLoad("M3", 7.68, 7.68, , , Transverse, "Case 1")
    Call myModel.AddMemberDistLoad("M2", -5.76, -5.76, , , Axial, "Case 1")
    Call myModel.AddMemberDistLoad("M3", 5.76, 5.76, , , Axial, "Case 1")
    
    'Add a load combination
    Call myModel.AddLoadCombo("Combo 1", Array("Case 1"), Array(1))
    
    'Clear old results
    Sheet1.Cells.Clear
    
    'Print the max/min bending moments
    Range("A1").Value = "Bending Moments"
    Range("A2").Value = "Member ID"
    Range("B2").Value = "Mmax"
    Range("C2").Value = "Mmin"
    Range("A3").Value = "M2"
    Range("B3").Value = myModel.GetMaxMoment("M2", "Combo 1")
    Range("C3").Value = myModel.GetMinMoment("M2", "Combo 1")
    Range("A4").Value = "M3"
    Range("B4").Value = myModel.GetMaxMoment("M3", "Combo 1")
    Range("C4").Value = myModel.GetMinMoment("M3", "Combo 1")
    Range("A5").Value = "M4"
    Range("B5").Value = myModel.GetMaxMoment("M4", "Combo 1")
    Range("C5").Value = myModel.GetMinMoment("M4", "Combo 1")
    Range("D5").Value = myModel.GetMaxShear("M4", "Combo 1")
    Range("E5").Value = myModel.GetMinShear("M4", "Combo 1")
    
    'Print force diagrams
    Range("A7").Value = "M2 Moment Diagram"
    Call myModel.GetMomentDiagram("M2", "Combo 1").PrintEZArray(Range("A8"))
    Range("D7").Value = "M3 Shear Diagram"
    Call myModel.GetShearDiagram("M3", "Combo 1").PrintEZArray(Range("D8"))
    Range("G7").Value = "M2 Axial Diagram"
    Call myModel.GetAxialDiagram("M2", "Combo 1").PrintEZArray(Range("G8"))
    
End Sub

'Finite Element Method, Example 4.10, Pages 196-198
'This example tests the program's ability to use single end releases, and to evaluate member loads
'This test was successfully run on 8/20/2017
Private Sub FEM_Problem_5_26()
    
    'Start a new finite element model
    Dim myModel As New FEModel
    
    'Define the nodes
    Call myModel.AddNode("N1", 8 * 12, 0)
    Call myModel.AddNode("N2", 8 * 12, 12 * 12)
    Call myModel.AddNode("N3", 0, 12 * 12)
    Call myModel.AddNode("N4", 20 * 12, 0)
    Call myModel.AddNode("N5", 20 * 12, 12 * 12)
    Call myModel.AddNode("N6", 28 * 12, 12 * 12)
    
    'Define the members
    Call myModel.AddMember("M1", "N1", "N2", 30000000, 300, 15)
    Call myModel.AddMember("M2", "N4", "N5", 30000000, 300, 15)
    Call myModel.AddMember("M3", "N3", "N2", 30000000, 600, 30)
    Call myModel.AddMember("M4", "N2", "N5", 30000000, 600, 30)
    Call myModel.AddMember("M5", "N5", "N6", 30000000, 600, 30)
    
    'Define the supports
    Call myModel.EditSupport("N1", True, True, False)
    Call myModel.EditSupport("N4", True, True, True)
    Call myModel.EditSupport("N3", True, True, True)
    Call myModel.EditSupport("N6", True, True, True)
    
    'Define the member end releases
    Call myModel.EditEndReleases("M1", False, False)
    Call myModel.EditEndReleases("M2", True, False)
    
    'Add member loads
    Call myModel.AddMemberDistLoad("M3", 1000 / 12, 1000 / 12)
    Call myModel.AddMemberDistLoad("M4", 1000 / 12, 1000 / 12)
    Call myModel.AddMemberDistLoad("M5", 1000 / 12, 1000 / 12)
    
    'Add a load combination
    Call myModel.AddLoadCombo("Combo 1", Array("Case 1", "Case 2"), Array(1, 1.2))
    
    'Clear old results
    Sheet1.Cells.Clear
    
    'Support Reactions
    Sheet1.Range("A1") = "Node N1 and N4 Reactions"
    Sheet1.Range("A2") = myModel.GetReaction("N1", FX, "Combo 1")
    Sheet1.Range("A3") = myModel.GetReaction("N1", FY, "Combo 1")
    Sheet1.Range("A4") = myModel.GetReaction("N1", MZ, "Combo 1")
    Sheet1.Range("B2") = myModel.GetReaction("N4", FX, "Combo 1")
    Sheet1.Range("B3") = myModel.GetReaction("N4", FY, "Combo 1")
    Sheet1.Range("B4") = myModel.GetReaction("N4", MZ, "Combo 1")
    
    'Joint Displacements
    Sheet1.Range("A6") = "Node N2 Displacements"
    Sheet1.Range("A7") = myModel.GetDisp("N2", DX, "Combo 1")
    Sheet1.Range("A8") = myModel.GetDisp("N2", DY, "Combo 1")
    Sheet1.Range("A9") = myModel.GetDisp("N2", RZ, "Combo 1")
    
End Sub

'Structural Analysis, Example 6.4, Pages 242-245
'This example tests the following:
    '- The program's ability to correctly calculate deflections at any point on a member
    '- The program's ability to correctly calculate the slope at any point on a member
    '  from which the deflections were derived.
'This test was successfully run on 10/19/2017
Private Sub SA_Example_6_4()
    
    'Start a new finite element model
    Dim myModel As New FEModel
    
    'Define the nodes
    Call myModel.AddNode("A", 0, 0)
    Call myModel.AddNode("D", 40 * 12, 0)
    
    'Define the members
    Call myModel.AddMember("M1", "A", "D", 1800, 46000, 25)
    
    'Define the supports
    Call myModel.EditSupport("A", True, True, False)
    Call myModel.EditSupport("D", False, True, False)
    
    'Add member point loads
    Call myModel.AddMemberPointLoad("M1", 60, 20 * 12, Transverse)
    Call myModel.AddMemberPointLoad("M1", 40, 30 * 12, Transverse)
    
    'Add a load combination
    Call myModel.AddLoadCombo("Combo 1", Array("Case 1"), Array(1))
    
    'Clear old results
    Sheet1.Cells.Clear
    
    'Print member displacements
    Sheet1.Range("A1") = "Member Displacements"
    Sheet1.Range("A2") = myModel.GetMemberDisp("M1", 20 * 12, "Combo 1")
    Sheet1.Range("A3") = myModel.GetMemberDisp("M1", 30 * 12, "Combo 1")
    
End Sub

'Use and change this procedure to isolate problems in the code
Private Sub DebugProgram()
    
    Dim myModel As New FEModel
    
    Call myModel.AddNode("N1", 0, 0)
    Call myModel.AddNode("N2", 10, 0)
    
    Call myModel.AddMember("M1", "N1", "N2", 29000 * 12 ^ 2, 200 / 12 ^ 4, 10 / 12 ^ 2)
    
    Call myModel.AddMemberMoment("M1", 20, 5)
    Call myModel.EditSupport("N1", True, True, True)
    'Call myModel.EditSupport("N2", True, True, True)
    
    'Add a load combination
    Call myModel.AddLoadCombo("Combo 1", Array("Case 1"), Array(1))
    
    Sheet1.Cells.Clear
    
    Range("A1").Value = myModel.GetMinMoment("M1")
    
    Range("A7").Value = "M1 Moment Diagram"
    Call myModel.GetMomentDiagram("M1").PrintEZArray(Range("A8"))
    
    Range("D7").Value = "M1 Shear Diagram"
    Call myModel.GetShearDiagram("M1").PrintEZArray(Range("D8"))
    
End Sub
