# 2D-Frame-Analysis---VBA

## Getting Started
This project gives you the ability to perform 2D structural analysis on frame/truss/beam structures using Visual Basic for Applications within Microsoft Excel. To get started, download and import the '.bas' and '.cls' files into your VBA project tree in the Visual Basic Editor built into Excel.

A FEModel class module has been provided to do all the heavy lifting of coordinating between the other modules. This is the only class you need to worry about to access to build, analyze, and get results from your finite element model. You can start by instantiating a new instance of the FEModel class using a statement such as:

    'Beginning a new 2D finite element model
    Dim myModel as New FEModel

From there you can access all the functions and properties you need to run a 2D finite element model using the '.' operator with your newly instantiated class. For example:

    'Defining a node
    Call myModel.AddNode("N1", 0, 0)

The VBA editor's intellisense will guide you along.

## More Documentation to Come
I plan to improve the documentation in the future. If you need further help you can run the subroutines in the "TestRoutines" module. These are textbook examples I have run to validate the code is executing correctly.

As of now, the FEModel class is very unforgiving if you have any instabilities in your model. It will stop code execution and give you an error. A common error people make is to have too many end releases at a node, allowing the node to spin.

## Help Wanted
I would love to have help streamlining this project. Let me know if you are capable and interested in making a contribution. I am a structural engineer first and a hobbyist programmer second, so it has taken years to learn how to bring it to this point. Enjoy!
