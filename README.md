All-Branch PAML
==========

An automated pipeline for selection test with designated model sets on each and every branch of a tree.

**Version**: 0.1.1

**Authors**: Qiyun Zhu, Katharina Dittmar

**Affiliation**: Department of Biological Sciences, University at Buffalo, State University of New York, Buffalo, USA

**License**: [BSD 2-clause](http://opensource.org/licenses/BSD-2-Clause).

#### Functionality

This Program calls the [PAML](http://abacus.gene.ucl.ac.uk/software/paml.html) package to test selection on each and every branch of a phylogenetic tree. It runs the `codeml` program with two (or more) alternative model sets (`codeml.ctl`), and performs likelihood ratio test (LRT) to assess the significance of the difference among likelihoods. The resulting p-values as well as the foreground omega (dN/dS) values are reported in tabular form, and also added to the tree as branch labels for further visualization.

The typical usage of this program is to test a branch model (two-ratio vs. one-ratio) or a branch-site model (model A vs. A1) (see PAML documents for details). The trick is that the user *does not have to manually insert a `#1` label into the input tree* file at the branch they would like to test, but intead, the program will do this automatically for all branches.

In such case, the first model set is the null model, assuming a universal omega value for the whole tree, and the second model set is the alternative model, with a different omega value at one branch. The program runs the null model for one time and the alternative model for *n* times (*n* = number of branches), and performs *n* LRTs.

The program provides a graphic user interface (GUI) to create and edit model sets (`codeml.ctl`), which can be imported / exported, in connection with existing PAML analysis configurations. It also provides an interface to display the result in tabular format and in Newick tree format, which can be further visualized in external tree-viewing programs.

#### Notes

**System requirements**: The program runs in all versions of 32/64-bit Windows. It was written in Visual Basic 6.0 (VB6). If you encounter a error message saying “missing files…”, you may want to download and install VB6 run-time files from the Microsoft [website](http://support2.microsoft.com/vbruntime).

**Input files**: A sequence alignment file and a tree file that are readable by PAML may be used. For best compatibility, the tree may be in Newick format, with or without branch lengths. The PAML package component `codeml.exe` should be present.

**Output files**: The resulting tables and trees are saved as `report.txt`, with trees further saved in Newick format. Depending on user choice, the original codeml output file `mlc` is saved after each run and renamed as `modelIDbranchID.txt`.

Please visit the Dittmar Lab [website](http://katharina-dittmar.squarespace.com/) and Qiyun's [blog](http://qiyunzhu.blogspot.com/) for details.

Please contact Qiyun Zhu (<qiyunzhu@gmail.com>) for any questions.
