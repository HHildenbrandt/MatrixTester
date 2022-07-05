# MatrixTester Excel Add-In

MatrixTester is developed for testing and describing social interactions in a group of individuals in Microsoft Excel on Windows. For testing it comprises the TauKr test, among others. This can be used for investigating whether individuals in a group direct more social acts (e.g. grooming) to those to whom they are more closely related, to whom they also direct other behavioural acts more frequently (e.g. support in fights, aggression) or from whom they receive more social acts in return (as happens relative reciprocity and interchange) (Hemelrijk 1990a). It can also test whether individuals direct acts more to those partners the higher their rank.  
It includes a partial TauKr test, by which a social variable may be controlled for. For instance, for studying whether grooming is reciprocated at a group level when excluding the effects of dominance rank of the interaction partner (Hemelrijk 1990b).
As a robust measure of the dominance hierarchy it comprises the average dominance index, which represents the average winning tendency of an individual with all its (sub)group members (Hemelrijk et al. 2005). It also comprises a measure of the degree of dominance of one type of individuals over another, here illustrated with degree of dominance of females over males (Hemelrijk et al. 2003, Hemelrijk et al. 2008).
As a measure of the diversity of an individual’s social interaction partners, it comprises the Berger Parker Dominance Index (Southwood 1978).

This program is developed to deal with the problem of statistical dependency. In a group of individuals each individual may  interact socially with all group members, as both as an actor and a receiver of acts. Thus, each individual is present in the data several times, as a potential actor with all group members and as a potential receiver.  It is present in a row and in a column, see figure below. This recurrence causes statistical dependency and thus, special statistical methods need to be used dealing with this dependency (Hubert 1987). A special permutation procedure deals with this dependency by reshuffling the individuals in rows and columns simultaneously (if identical individuals are present in rows and columns) and reshuffles individuals in rows and columns alternatingly, if they belong to two different categories (Hemelrijk 1990a). Individuals belong to two categories for instance in the case they differ in sex (males and females),  in age (youngsters and adults) or in color (red and blue cichlids), etcetera.  

Please always acknowledge usage of the program and method in your paper and send Hemelrijk a reprint, c.k.hemelrijk@rug.nl. I can then discuss in future papers all kinds of usage and develop further extensions.

Charlotte K Hemelrijk,University of Groningen.  
Contact: c.k.hemelrijk@rug.nl

## Choosing the matching version

Unzip the MatrixTester Add-In that matches your version of Excel. Use `MatrixTester-AddIn.zip` if your Office installation is 32Bit or `MatrixTester-AddIn64.zip` if your Office installation is 64Bit. You can find this information (32Bit or 64Bit) under *Properties* or *Help* in Excel.

### Install the MatrixTester Add-In (Office 2016 and later)

1. Click the *File* tab in Excel, click *Options*, and then click the *Add-ins* category.
2. In the *Manage* box, click *Excel Add-ins*, and then click *Go*. The *Add-ins dialog* box appears.
3. In the *Add-ins available box*, browse to the folder you unzipped MatrixTester and select the file that ends with `.xll`.
4. Check the check box next to the Add-in that you want to activate, and then click *OK*.

### Uninstall the MatrixTester Add-In (Office 2016 and later)

1. Delete the MatrixTester installation folder.
2. Start Excel anew.
3. A warning appears that the Add-In cannot be found. Click *OK*.
4. Goto *File*, *Options*, *Add-ins*, click *Go*.
5. In the Add-In list that appear, un-check MatrixTester.
6. A message appears that it is deleted. Click *OK*.

### Update or re-install the MatrixTester Add-In (Office 2016 and later)

If you want to update or re-install the MatrixTester Add-In, make sure you un-install the previous Add-In.

### Troubleshooting

Microsoft has frequently changed the handling of Add-Ins. You should be able to find instructions in the documentation of your Office version thought. Another source of trouble is security - your institution might block you from installing *any* software or may have policies in place that restricts the use of Add-Ins. Please contact your administrator to help you out.
Unfortunately, we currently don't support *Office for MacOS* due to a lack of resources.

# Usage

## Reading in matrices

Before reading in matrices, first make sure that Excel uses a dot for a decimal point. You do this by going in the control Panel, under Region and language and then under input, format and other settings. Here you adjust such that decimal settings are indicated by a dot.

To read in the legacy matrices used with former versions of this program, **TYPIN**, **MATSQUAR**, **MATRECT** and **MATRIXTESTER**:

1. goto *add-in*.
2. select *options*.
3. select *legacy Matrices*.

To read in .csv matrices

1. goto *Data*.
2. click on *fromText*.
3. import the file in Excel as usual.

## Manipulations with matrices

**To make a matrix symmetrical (sum the values in the top and bottom triangle of the matrix)**

1. click on an empty cell.
2. type `'=mt'`.
3. go down with the cursor till `Symm`.
4. press the tab-button.
5. select the matrix with names of rows and columns.
6. close with *return*.

**To sort rows and columns of a matrix (if individuals in rows and columns are identical)**

1. click on an empty cell.
2. type `'=mt'`
3. go down with the cursor till `Sort.Full`.
4. press the *tab-button*.
5. select matrix.
6. select column for sorting.
7. press *shift*, *ctrl*, *return*.

As an example, try to sort the following matrix by `gender`:

```
        m1	f1	m2	f2	f3	gender
m1	0	6	9	8	5	m
f1	0	0	4	6	0	f
m2	0	2	0	4	7	m
f2	1	0	5	0	3	f
f3	0	0	2	3	0	f
```
The result should be:
```
	f1	f2	f3	m1	m2	gender
f1	0	6	0	0	4	f
f2	0	0	3	1	5	f
f3	0	3	0	0	2	f
m1	6	8	5	0	9	m
m2	2	4	7	0	0	m
```

**To sort rows of a matrix (if individuals in rows and columns differ)**

1. click on an empty cell.
2. type `'=mt'`
3. go down with the cursor till `Sort.Rows`.
4. press the *tab-button*.
5. select matrix.
6. select column for sorting.
7. press *shift*, *ctrl*, *return*.

## Descriptive statistics

**To calculate the average dominance index**, which represents the average winning tendency of an individual with all its (sub)group members (Hemelrijk et al. 2005).

1. click on an empty cell.
2. type `'=mt'`.
3. go down with the cursor till `avgDI`.
4. select the matrix
5. press *return*

**To calculate the degree of female dominance** as the summed number of males ranking below all females divided by the sum of the maximum potential number of males ranking below all females  males (Hemelrijk et al. 2003, Hemelrijk et al. 2008).

1. click on an empty cell.
2. type `'=mt'`.
3. go down with the cursor till `FemDom`.
4. press the *tab-button*.
5. select the matrix with names of rows and columns.
6. type `','` (comma). Sometimes your computer is expecting `';'` if comma does not work adjust your Excel setting via Control panel).
7. To select the names of the females, chose female-name by *ctrl click* or chose a row of connected names, type `','`, chose another female by *ctrl click* and so on.
8. execute the calculation by pressing *shift*, *ctrl*, *return*.

**To calculate the Berger Parker Dominance Index (Southwood 1978)**, which describes the diversity of an individual’s social interaction partners

1. click on an empty cell.
2. type `'=mt'`.
3. go down with the cursor till `BPI`.
4. press the *tab*-button.
5. select the matrix with names of rows and columns.
6. close by pressing *return*

## Different statistical tests

As a default, the TauKr test is given rather than the Kr-value, because the TauKr value corrects for the number of individuals and ties, whereas Kr does not  (Hemelrijk 1990b).  The Kr test and other tests, such as the K test, Kc, R and Z test are  all given under  MatrixTester options, because they suffer from the disadvantage of being uncorrected to group size.

## Actual testing

Matrix-correlations can be executed either between individuals of one category, where the individuals in the rows and columns are identical, or between individuals of two categories, where individuals of rows and columns differ in type and none of the individuals is found in both rows and columns. The number of permutations is per default 10.000, but it can be adjusted under MatrixTester options. If in case of two categories the same individual is found in row and column, an error message will be given and the individual has to be omitted from either the row or the column.

**To execute the TauKr test**

1. click on an empty cell.
2. type `'=mt'`.
3. go down with the cursor till `TauKr`.
4. press the *tab*-button.
5. select the matrix with names of rows and colums.
6. type `','` (comma) (if it does not work adjust your Excel setting via Control panel).
7. select the second matrix.
8. press *return*.

**To execute the partial TauKr test**

1. click on an empty cell.
2. type `'=mt'`.
3. go down with the cursor till `TauKr`.
4. press the *tab*-button.
5. select the matrix with names of rows and colums.
6. type `','` (comma) (if it does not work adjust your Excel setting via Control panel).
7. select the second matrix.
8. type type `','` (comma) (if it does not work adjust your Excel setting via Control panel).
9. select the third matrix (to be controlled for).
10. execute by clicking *ctrl*, *shift*, *return*.

**To execute the K test, KcTest, RTest and Ztest**

1. Goto *MatrixTester*, *options* and select *deprecated functions*, start Excel anew.
2. click on an empty cell.
3. type `'=mt'`
4. go down with the cursor till you find the test `(K, Kc, R, Z)`.
5. press the *tab*-button.
6. select the matrix with names of rows and colums.
7. type type `','` (comma) (if it does not work adjust your Excel setting via Control panel).
8. select the second matrix.
9. press *return*.

## Dominance Rank

You may want to order individuals according to rank in a matrix. For this you can use the weighted dominance index (the average of the ratio of winning per individuals), this equals the David Score, but is calculated in a more simple and straightforward way taking care of missing values (Hemelrijk et al. 2005). 

**To order individuals in a matrix according to rank**

1. click on an empty cell.
2. type `'=mt'`.
3. go down with the cursor till `AvgDI`. 
4. press the *tab*-button. 
5. select the matrix with names of rows and colums. 
6. press *return*.

## References

* Hemelrijk CK, Wantia J, Isler K (2008) Female dominance over males in primates: Self-organisation and sexual dimorphism. PLoS ONE 3:e2678
* Hemelrijk CK, Wantia J, Daetwyler M (2003) Female co-dominance in a virtual world: Ecological, cognitive, social and sexual causes. Behaviour 140:1247-1273
* Hemelrijk CK, Wantia J, Gygax L (2005) The construction of dominance order: comparing performance of five different methods using an individual-based model. Behaviour 142:1043-1064
* Hemelrijk CK (1990a) Models of, and tests for, reciprocity, unidirectional and other social interaction patterns at a group level. Anim Behav 39:1013-1029
* Hemelrijk CK (1990b) A Matrix Partial Correlation Test used in Investigations of Reciprocity and Other Social Interaction Patterns at Group Level. J. theor. Biol. 143:405-420
* Hubert LJ (1987) Assignment methods in combinatorial data analysis. Marcel Dekker, Inc., New York
* Southwood TRE (1978) Ecological methods with particular reference to the study of insect populations. Chapman and Hall, London and New York
