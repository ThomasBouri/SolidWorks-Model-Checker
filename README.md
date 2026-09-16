# SolidWorks-Model-Checker

A SolidWorks macro used to test and grade how a part was built and evaluates aspects that SolidWorks does not.
A tool specifically designed to help students who have taken CAD courses but could have developed bad habits.
The part 

This is scored under two categories; Robustness and Conventions.

## Robustness 
_Aspects of a part that could have mechanical consequences_
- Sketches Fully Defined

## Conventions
_Departures from practice that can have an impact on the next person to open the file_
- Features renamed
- No unused sketches
- Material Assigned
- File properties filled

## Caution
A third section that is reported but not scored. Sketches anchored on model geometry rather than reference planes _could_ cause issues (if the geometry they reference moves or is consumed). As such, this section of the rubric is only used to remind the user of how many sketches are on faces.

# How It Works
This process of creating the rubric is defined with 3 **distinct** sections in mind.

### Scan
Reads the model and records raw data. No opinions and no scoring.

### Judge
Reads what scan recorded, and does not interact with SolidWorks. Applies the weights, computes the scores for each section, and builds the findings list.

### Present
Reads what Judge produced and presents the rubric. Does not interact with the Scan section, nor SolidWorks.


## Installation & Use
1. In SolidWorks: **Tools -> Macro -> New** (save as ModelCheck.swp)
2. In the VBA Editor: **File -> Import File** (choose the downloaded ModelCheck.bas)
3. Delete the empty default Module1
4. Run `Main` and follow the pop-up prompts within SolidWorks
