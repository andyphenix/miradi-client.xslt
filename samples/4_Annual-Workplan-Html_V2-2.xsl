<!-- Copyright 2012, Bush Heritage Australia, Melbourne, Australia. -->
<!-- This XSL style-sheet is free software. You can redistribute it and/or modify it under the -->
<!-- terms of the GNU General Public License version 3, as published by the Free Software Foundation.-->
<!-- The style-sheet is distributed in the hope that it will be useful, but without any warranty; -->
<!-- without even the implied warranty of merchantability or fitness for a particular purpose.  -->
<!-- See the GNU General Public License for more details http://www.gnu.org/licenses. -->
<xsl:stylesheet 
	version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:o="urn:schemas-microsoft-com:office:office" 
	xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" 
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:v="urn:schemas-microsoft-com:vml">

	<xsl:param name="FinancialYear">2012</xsl:param>

	<xsl:template match="ConservationProject">
		<html>
			<head>
				<style>
					body
					{
					font-family: Arial;
					font-size: 10pt
					}

					.activitytitle
					{
					font-size: 10pt;
					font-weight: bold;
					}

					.activity
					{
					font-size: 10pt;
					}

					.objective
					{
					color: #4F81BD;
					font-family: Arial;
					font-size: 10pt;
					}

					.title
					{
					font-family: Calibri;
					font-size: 18pt;
					}

					.resultschaintitle
					{
					background-color: #E0E0FF;
					color: #4F81BD;
					font-family: Arial;
					font-size: 14pt;
					}

					.strategyid
					{
					color: #4F81BD;
					font-family: Arial;
					font-size: 13pt;
					mso-number-format:\@
					}

					.strategytitle
					{
					color: #4F81BD;
					font-family: Arial;
					font-size: 13pt;
					}

					.bordernonumberformat
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 0.5pt;
					border-color: gray;
					font-family: Arial;
					font-size: 10pt;
					vertical-align: top;
					};

					.border
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 0.5pt;
					border-color: gray;
					font-family: Arial;
					font-size: 10pt;
					vertical-align: top;
					mso-number-format:\#\,\#\#0\;
					};

					.borderresultschain
					{
					background-color: #E0E0FF;
					border-collapse: collapse;
					border-style: solid;
					border-width: 0.5pt;
					border-color: gray;
					font-family: Arial;
					font-size: 10pt;
					vertical-align: top;
					mso-number-format:\#\,\#\#0\;
					}

					.borderresultschainheader
					{
					background-color: #E0E0FF;
					border-collapse: collapse;
					border-style: solid;
					border-width: 0.5pt;
					border-color: gray;
					font-weight: bold;
					font-family: Arial;
					font-size: 10pt;
					vertical-align: top;
					}

					.borderstrategy
					{
					background-color: #E0E0E0;
					border-collapse: collapse;
					border-style: solid;
					border-width: 0.5pt;
					border-color: gray;
					font-family: Arial;
					font-size: 10pt;
					vertical-align: top;
					mso-number-format:\#\,\#\#0\;
					};

					.borderstrategyheader
					{
					background-color: #E0E0E0;
					border-collapse: collapse;
					border-style: solid;
					border-width: 0.5pt;
					border-color: gray;
					font-weight: bold;
					font-family: Arial;
					font-size: 10pt;
					vertical-align: top;
					}

					.projectcode
					{
					font-family: Calibri;
					font-size: 18pt;
					mso-number-format:\@;
					}

				</style>
			</head>
			<body>
				<table>
					<xsl:apply-templates select="ProjectSummary" mode="title" />
					<xsl:apply-templates select="ResultsChainPool" mode="resultschain"/>
				</table>
			</body>
		</html>
	</xsl:template>

	<xsl:template match="ProjectSummary" mode="title">
		<tr>
			<td colspan="3" class="title">
				Annual Workplan and Budget
			</td>
			<td colspan="12" class="title">
				<xsl:value-of select="ProjectSummaryProjectName"/><xsl:text>&#x20;</xsl:text><xsl:value-of select="$FinancialYear"/>/<xsl:value-of select="$FinancialYear+1"/>
			</td>
			<td colspan="5" class="title">
				Project Code:<xsl:text>&#x20;</xsl:text><xsl:value-of select="ProjectSummaryOtherOrgProjectNumber"/>
			</td>
		</tr>
		<tr>
			<td></td>
		</tr>
	</xsl:template>

	<xsl:template name="rowcount">
		<xsl:param name="rc" />
		<xsl:param name="TaskId" />
		<xsl:variable name="result">
			<xsl:value-of select="$rc + 4 + count(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = /ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 6)]/../../..) + count(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = /ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../..) + count(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = /ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 4)]/../../..) + count(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = /ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../..)"/>
		</xsl:variable>
		<xsl:value-of select="$result"/>
	</xsl:template>

	<xsl:template match="ResultsChainPool" mode="resultschain">
		<xsl:for-each select="ResultsChain">
			<xsl:sort select="ResultsChainId" />

			<xsl:variable name="DiagramFactorId" select="ResultsChainDiagramFactorIds/DiagramFactorId" />

			<!-- Calculate the total number of rows across all strategies for this results chain -->
			<xsl:variable name="ResultsChainRows">
				<xsl:variable name="StrategyId" select="/ConservationProject/DiagramFactorPool/DiagramFactor[@Id = $DiagramFactorId]/DiagramFactorWrappedFactorId/WrappedByDiagramFactorId/StrategyId" />
				<xsl:variable name="ActivityId" select="/ConservationProject/StrategyPool/Strategy[@Id = $StrategyId]/StrategySortedActivityIds/ActivityId" />
				<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $ActivityId]/TaskAssignmentIds/ResourceAssignmentId" />
				<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $ActivityId]/TaskExpenseIds/ExpenseAssignmentId" />
				<xsl:variable name="DaysMonthsRows" select="count(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 4)]/../../..)" />
				<xsl:variable name="DaysYearsRows" select="count(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../..)" />
				<xsl:variable name="ExpensesMonthsRows" select="count(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 4)]/../../..)" />
				<xsl:variable name="ExpensesYearsRows" select="count(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../..)" />
				<xsl:value-of select="1 + count($StrategyId)*7 + count($ActivityId)*3 + $DaysMonthsRows + $DaysYearsRows + $ExpensesMonthsRows + $ExpensesYearsRows"/>
			</xsl:variable>

			<tr>
				<td colspan="2" class="resultschaintitle">
					Results Chain:
				</td>
				<td colspan="18" class="resultschaintitle">
					<xsl:value-of select="ResultsChainName"/>
				</td>
			</tr>
			<tr>
				<td class="borderresultschainheader"></td>
				<td class="borderresultschainheader">Period</td>
				<td class="borderresultschainheader">Details</td>
				<td class="borderresultschainheader">Acct Code</td>
				<td class="borderresultschainheader">Budget 1</td>
				<td class="borderresultschainheader">Budget 2</td>
				<td class="borderresultschainheader">Funding</td>
				<td class="borderresultschainheader">Apr</td>
				<td class="borderresultschainheader">May</td>
				<td class="borderresultschainheader">Jun</td>
				<td class="borderresultschainheader">Jul</td>
				<td class="borderresultschainheader">Aug</td>
				<td class="borderresultschainheader">Sep</td>
				<td class="borderresultschainheader">Oct</td>
				<td class="borderresultschainheader">Nov</td>
				<td class="borderresultschainheader">Dec</td>
				<td class="borderresultschainheader">Jan</td>
				<td class="borderresultschainheader">Feb</td>
				<td class="borderresultschainheader">Mar</td>
				<td class="borderresultschainheader">TOTAL</td>
			</tr>
			<tr>
				<td class="borderresultschainheader">Days</td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[3]C1:R[<xsl:value-of select="$ResultsChainRows + 2" />]C1,"Days",R[3]C)/2
				</td>
				<td class="borderresultschain">
					=SUM(RC[-12]:RC[-1])
				</td>
			</tr>
			<tr>
				<td class="borderresultschainheader">Expenses</td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[2]C1:R[<xsl:value-of select="$ResultsChainRows + 1" />]C1,"Expenses",R[2]C)/2
				</td>
				<td class="borderresultschain">
					=SUM(RC[-12]:RC[-1])
				</td>
			</tr>
			<tr>
				<td class="borderresultschainheader">Total</td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain"></td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUMIF(R[1]C1:R[<xsl:value-of select="$ResultsChainRows" />]C1,"Total",R[1]C)/2
				</td>
				<td class="borderresultschain">
					=SUM(RC[-12]:RC[-1])
				</td>
			</tr>

			<tr>
				<td></td>
			</tr>

			<xsl:variable name="StrategyId" select="/ConservationProject/DiagramFactorPool/DiagramFactor[@Id = $DiagramFactorId]/DiagramFactorWrappedFactorId/WrappedByDiagramFactorId/StrategyId" />

			<!-- Prepare report for each strategy -->

			<xsl:for-each select="/ConservationProject/StrategyPool/Strategy[@Id = $StrategyId]">

				<xsl:variable name="ObjectiveId" select="StrategyObjectiveIds/ObjectiveId" />

				<!-- Calculate the total number of rows across all activities for this strategy -->

				<xsl:variable name="StrategyRows">
					<xsl:variable name="ActivityId" select="StrategySortedActivityIds/ActivityId" />
					<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $ActivityId]/TaskAssignmentIds/ResourceAssignmentId" />
					<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $ActivityId]/TaskExpenseIds/ExpenseAssignmentId" />
					<xsl:variable name="DaysMonthsRows" select="count(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 4)]/../../..)" />
					<xsl:variable name="DaysYearsRows" select="count(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../..)" />
					<xsl:variable name="ExpensesMonthsRows" select="count(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 4)]/../../..)" />
					<xsl:variable name="ExpensesYearsRows" select="count(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../..)" />
					<xsl:value-of select="1 + count($ActivityId)*3 + $DaysMonthsRows + $DaysYearsRows + $ExpensesMonthsRows + $ExpensesYearsRows"/>
				</xsl:variable>
				
				<tr>
					<td class="strategytitle">Strategy:</td>
					<td colspan="15" class="strategytitle">
						<xsl:value-of select="StrategyName"/>
					</td>
					<td colspan="4" class="strategytitle">
						Activity Code: <xsl:value-of select="StrategyId"/>
					</td>
				</tr>
				<tr>
					<td></td>
				</tr>
				<xsl:for-each select="/ConservationProject/ObjectivePool/Objective[@Id = $ObjectiveId]/ObjectiveName">
					<tr>
						<td class="objective">
							Objective:
						</td>
						<td class="objective" colspan="33">
							<xsl:value-of select="text()"/>
						</td>
					</tr>
				</xsl:for-each>
				<tr>
					<td class="borderstrategyheader"></td>
					<td class="borderstrategyheader">Period</td>
					<td class="borderstrategyheader">Details</td>
					<td class="borderstrategyheader">Acct Code</td>
					<td class="borderstrategyheader">Budget 1</td>
					<td class="borderstrategyheader">Budget 2</td>
					<td class="borderstrategyheader">Funding</td>
					<td class="borderstrategyheader">Apr</td>
					<td class="borderstrategyheader">May</td>
					<td class="borderstrategyheader">Jun</td>
					<td class="borderstrategyheader">Jul</td>
					<td class="borderstrategyheader">Aug</td>
					<td class="borderstrategyheader">Sep</td>
					<td class="borderstrategyheader">Oct</td>
					<td class="borderstrategyheader">Nov</td>
					<td class="borderstrategyheader">Dec</td>
					<td class="borderstrategyheader">Jan</td>
					<td class="borderstrategyheader">Feb</td>
					<td class="borderstrategyheader">Mar</td>
					<td class="borderstrategyheader">TOTAL</td>
					<td></td>
					<td>Rate</td>
					<td>Apr</td>
					<td>May</td>
					<td>Jun</td>
					<td>Jul</td>
					<td>Aug</td>
					<td>Sep</td>
					<td>Oct</td>
					<td>Nov</td>
					<td>Dec</td>
					<td>Jan</td>
					<td>Feb</td>
					<td>Mar</td>
				</tr>
				<tr>
					<td class="borderstrategyheader">Days</td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[3]C1:R[<xsl:value-of select="$StrategyRows + 2" />]C1,"Days",R[3]C)
					</td>
					<td class="borderstrategy">
						=SUM(RC[-12]:RC[-1])
					</td>
				</tr>
				<tr>
					<td class="borderstrategyheader">Expenses</td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[2]C1:R[<xsl:value-of select="$StrategyRows + 1" />]C1,"Expenses",R[2]C)
					</td>
					<td class="borderstrategy">
						=SUM(RC[-12]:RC[-1])
					</td>
				</tr>
				<tr>
					<td class="borderstrategyheader">Total</td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy"></td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUMIF(R[1]C1:R[<xsl:value-of select="$StrategyRows" />]C1,"Total",R[1]C)
					</td>
					<td class="borderstrategy">
						=SUM(RC[-12]:RC[-1])
					</td>
				</tr>

				<!-- Within each strategy, report on each activity -->
				
				<xsl:for-each select="StrategySortedActivityIds/ActivityId">
					<xsl:variable name="TaskId">
						<xsl:value-of select="text()"/>
					</xsl:variable>

					<xsl:variable name="ThisStrategyId" select="../../StrategyId" />

					<tr>
						<td>
						</td>
					</tr>
					<tr>
						<td class="activitytitle">
							Activity
						</td>
						<td colspan="19" class="activity">
							<xsl:value-of select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskName"/>
						</td>
					</tr>
					
					<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId" />
					<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId" />
					<xsl:variable name="ExpenseAssignmentFundingSourceId" select="/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentFundingSourceId/FundingSourceId" />
					<xsl:variable name="ResourceAssignmentFundingSourceId" select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentFundingSourceId/FundingSourceId" />

					<!-- Calculate the number of detail rows for this activity. -->

					<xsl:variable name="Rows">
						<xsl:variable name="DaysMonthsRows" select="count(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 4)]/../../..)" />
						<xsl:variable name="DaysYearsRows" select="count(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../..)" />
						<xsl:variable name="ExpensesMonthsRows" select="count(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 4)]/../../..)" />
						<xsl:variable name="ExpensesYearsRows" select="count(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../..)" />
						<xsl:value-of select="$DaysMonthsRows + $DaysYearsRows + $ExpensesMonthsRows + $ExpensesYearsRows"/>
					</xsl:variable>

					<!-- Do day/month rows -->

					<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 4)]/../../..">
						<xsl:variable name="BudgetCategoryTwoId" select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentBudgetCategoryTwoId/BudgetCategoryTwoId" />
						<tr>
							<td class="border">
								Days
							</td>
							<td class="border">
								Month
							</td>
							<td class="border">
								<xsl:variable name="ResourceId" select="../ResourceAssignmentResourceId/ResourceId" />
								<xsl:value-of select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceResource_Id"/>
								<xsl:text>&#x20;</xsl:text>
								<xsl:value-of select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceGivenName"/>
								<xsl:text>&#x20;</xsl:text>
								<xsl:value-of select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceSurname"/>
							</td>
							<td class="border">
								<xsl:variable name="AccountingCodeId" select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentAccountingCodeId/AccountingCodeId" />
								<xsl:value-of select="/ConservationProject/AccountingCodePool/AccountingCode[@Id = $AccountingCodeId]/AccountingCodeCode"/>-<xsl:value-of select="$ThisStrategyId"/>-<xsl:value-of select="/ConservationProject/BudgetCategoryTwoPool/BudgetCategoryTwo[@Id = $BudgetCategoryTwoId]/BudgetCategoryTwoCode"/>
							</td>
							<td class="border">
								<xsl:variable name="BudgetCategoryOneId" select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentBudgetCategoryOneId/BudgetCategoryOneId" />
								<xsl:value-of select="/ConservationProject/BudgetCategoryOnePool/BudgetCategoryOne[@Id = $BudgetCategoryOneId]/BudgetCategoryOneCode"/>
							</td>
							<td class="border">
								<xsl:value-of select="/ConservationProject/BudgetCategoryTwoPool/BudgetCategoryTwo[@Id = $BudgetCategoryTwoId]/BudgetCategoryTwoCode"/>
							</td>
							<td class="bordernonumberformat">
								<xsl:value-of select="/ConservationProject/FundingSourcePool/FundingSource[@Id = $ResourceAssignmentFundingSourceId]/FundingSourceCode"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear and @Month = 4]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear and @Month = 5]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear and @Month = 6]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear and @Month = 7]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear and @Month = 8]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear and @Month = 9]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear and @Month = 10]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear and @Month = 11]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear and @Month = 12]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear+1 and @Month = 1]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear+1 and @Month = 2]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[@Year = $FinancialYear+1 and @Month = 3]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								=SUM(RC[-12]:RC[-1])
							</td>
							<td></td>
							<td>
								<xsl:value-of select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = /ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 4)]/../../../../ResourceAssignmentResourceId/ResourceId]/ProjectResourceDailyRate"/>
							</td>
							<td>=RC[-1]*RC[-15]</td>
							<td>=RC[-2]*RC[-15]</td>
							<td>=RC[-3]*RC[-15]</td>
							<td>=RC[-4]*RC[-15]</td>
							<td>=RC[-5]*RC[-15]</td>
							<td>=RC[-6]*RC[-15]</td>
							<td>=RC[-7]*RC[-15]</td>
							<td>=RC[-8]*RC[-15]</td>
							<td>=RC[-9]*RC[-15]</td>
							<td>=RC[-10]*RC[-15]</td>
							<td>=RC[-11]*RC[-15]</td>
							<td>=RC[-12]*RC[-15]</td>
						</tr>
					</xsl:for-each>

					<!-- Do day/year rows -->

					<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../..">
						<xsl:variable name="BudgetCategoryTwoId" select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentBudgetCategoryTwoId/BudgetCategoryTwoId" />
						<tr>
							<td class="border">
								Days
							</td>
							<td class="border">
								Year
							</td>
							<td class="border">
								<xsl:variable name="ResourceId" select="../ResourceAssignmentResourceId/ResourceId" />
								<xsl:value-of select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceResource_Id"/>
								<xsl:text>&#x20;</xsl:text>
								<xsl:value-of select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceGivenName"/>
								<xsl:text>&#x20;</xsl:text>
								<xsl:value-of select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceSurname"/>
							</td>
							<td class="border">
								<xsl:variable name="AccountingCodeId" select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentAccountingCodeId/AccountingCodeId" />
								<xsl:value-of select="/ConservationProject/AccountingCodePool/AccountingCode[@Id = $AccountingCodeId]/AccountingCodeCode"/>-<xsl:value-of select="$ThisStrategyId"/>-<xsl:value-of select="/ConservationProject/BudgetCategoryTwoPool/BudgetCategoryTwo[@Id = $BudgetCategoryTwoId]/BudgetCategoryTwoCode"/>
							</td>
							<td class="border">
								<xsl:variable name="BudgetCategoryOneId" select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentBudgetCategoryOneId/BudgetCategoryOneId" />
								<xsl:value-of select="/ConservationProject/BudgetCategoryOnePool/BudgetCategoryOne[@Id = $BudgetCategoryOneId]/BudgetCategoryOneCode"/>
							</td>
							<td class="border">
								<xsl:value-of select="/ConservationProject/BudgetCategoryTwoPool/BudgetCategoryTwo[@Id = $BudgetCategoryTwoId]/BudgetCategoryTwoCode"/>
							</td>
							<td class="bordernonumberformat">
								<xsl:value-of select="/ConservationProject/FundingSourcePool/FundingSource[@Id = $ResourceAssignmentFundingSourceId]/FundingSourceCode"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear and @StartMonth = 4]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear and @StartMonth = 5]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear and @StartMonth = 6]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear and @StartMonth = 7]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear and @StartMonth = 8]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear and @StartMonth = 9]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear and @StartMonth = 10]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear and @StartMonth = 11]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear and @StartMonth = 12]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear+1 and @StartMonth = 1]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear+1 and @StartMonth = 2]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[@StartYear = $FinancialYear+1 and @StartMonth = 3]/../../NumberOfUnits"/>
							</td>
							<td class="border">
								=SUM(RC[-12]:RC[-1])
							</td>
							<td></td>
							<td>
								<xsl:value-of select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = /ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../../../ResourceAssignmentResourceId/ResourceId]/ProjectResourceDailyRate"/>
							</td>
							<td>=RC[-1]*RC[-15]</td>
							<td>=RC[-2]*RC[-15]</td>
							<td>=RC[-3]*RC[-15]</td>
							<td>=RC[-4]*RC[-15]</td>
							<td>=RC[-5]*RC[-15]</td>
							<td>=RC[-6]*RC[-15]</td>
							<td>=RC[-7]*RC[-15]</td>
							<td>=RC[-8]*RC[-15]</td>
							<td>=RC[-9]*RC[-15]</td>
							<td>=RC[-10]*RC[-15]</td>
							<td>=RC[-11]*RC[-15]</td>
							<td>=RC[-12]*RC[-15]</td>
						</tr>
					</xsl:for-each>

					<!-- Expense/month assignment rows -->

					<xsl:for-each select="/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year = $FinancialYear and @Month &gt; 3) or (@Year = $FinancialYear+1 and @Month &lt; 4)]/../../..">
						<xsl:variable name="BudgetCategoryTwoId" select="../ExpenseAssignmentBudgetCategoryTwoId/BudgetCategoryTwoId" />
						<tr>
							<td class="border">
								Expenses
							</td>
							<td class="border">
								Month
							</td>
							<td class="border">
								<xsl:value-of select="../ExpenseAssignmentName"/>
							</td>
							<td class="border">
								<xsl:variable name="AccountingCodeId" select="../ExpenseAssignmentAccountingCodeId/AccountingCodeId" />
								<xsl:value-of select="/ConservationProject/AccountingCodePool/AccountingCode[@Id = $AccountingCodeId]/AccountingCodeCode"/>-<xsl:value-of select="$ThisStrategyId"/>-<xsl:value-of select="/ConservationProject/BudgetCategoryTwoPool/BudgetCategoryTwo[@Id = $BudgetCategoryTwoId]/BudgetCategoryTwoCode"/>
							</td>
							<td class="border">
								<xsl:variable name="BudgetCategoryOneId" select="../ExpenseAssignmentBudgetCategoryOneId/BudgetCategoryOneId" />
								<xsl:value-of select="/ConservationProject/BudgetCategoryOnePool/BudgetCategoryOne[@Id = $BudgetCategoryOneId]/BudgetCategoryOneCode"/>
							</td>
							<td class="border">
								<xsl:value-of select="/ConservationProject/BudgetCategoryTwoPool/BudgetCategoryTwo[@Id = $BudgetCategoryTwoId]/BudgetCategoryTwoCode"/>
							</td>
							<td class="bordernonumberformat">
								<xsl:value-of select="/ConservationProject/FundingSourcePool/FundingSource[@Id = $ExpenseAssignmentFundingSourceId]/FundingSourceCode"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear and @Month = 4]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear and @Month = 5]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear and @Month = 6]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear and @Month = 7]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear and @Month = 8]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear and @Month = 9]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear and @Month = 10]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear and @Month = 11]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear and @Month = 12]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear+1 and @Month = 1]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear+1 and @Month = 2]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesMonth[@Year = $FinancialYear+1 and @Month = 3]/../../Expense"/>
							</td>
							<td class="border">
								=SUM(RC[-12]:RC[-1])
							</td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
						</tr>
					</xsl:for-each>

					<!-- Expense/year assignment rows -->
					
					<xsl:for-each select="/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear = $FinancialYear and @StartMonth &gt; 3) or (@StartYear = $FinancialYear+1 and @StartMonth &lt; 4)]/../../..">
						<xsl:variable name="BudgetCategoryTwoId" select="../ExpenseAssignmentBudgetCategoryTwoId/BudgetCategoryTwoId" />
						<tr>
							<td class="border">
								Expenses
							</td>
							<td class="border">
								Year
							</td>
							<td class="border">
								<xsl:value-of select="../ExpenseAssignmentName"/>
							</td>
							<td class="border">
								<xsl:variable name="AccountingCodeId" select="../ExpenseAssignmentAccountingCodeId/AccountingCodeId" />
								<xsl:value-of select="/ConservationProject/AccountingCodePool/AccountingCode[@Id = $AccountingCodeId]/AccountingCodeCode"/>-<xsl:value-of select="$ThisStrategyId"/>-<xsl:value-of select="/ConservationProject/BudgetCategoryTwoPool/BudgetCategoryTwo[@Id = $BudgetCategoryTwoId]/BudgetCategoryTwoCode"/>
							</td>
							<td class="border">
								<xsl:variable name="BudgetCategoryOneId" select="../ExpenseAssignmentBudgetCategoryOneId/BudgetCategoryOneId" />
								<xsl:value-of select="/ConservationProject/BudgetCategoryOnePool/BudgetCategoryOne[@Id = $BudgetCategoryOneId]/BudgetCategoryOneCode"/>
							</td>
							<td class="border">
								<xsl:value-of select="/ConservationProject/BudgetCategoryTwoPool/BudgetCategoryTwo[@Id = $BudgetCategoryTwoId]/BudgetCategoryTwoCode"/>
							</td>
							<td class="bordernonumberformat">
								<xsl:value-of select="/ConservationProject/FundingSourcePool/FundingSource[@Id = $ExpenseAssignmentFundingSourceId]/FundingSourceCode"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear and @StartMonth = 4]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear and @StartMonth = 5]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear and @StartMonth = 6]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear and @StartMonth = 7]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear and @StartMonth = 8]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear and @StartMonth = 9]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear and @StartMonth = 10]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear and @StartMonth = 11]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear and @StartMonth = 12]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear+1 and @StartMonth = 1]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear+1 and @StartMonth = 2]/../../Expense"/>
							</td>
							<td class="border">
								<xsl:value-of select="DateUnitExpense/ExpensesDateUnit/ExpensesYear[@StartYear = $FinancialYear+1 and @StartMonth = 3]/../../Expense"/>
							</td>
							<td class="border">
								=SUM(RC[-12]:RC[-1])
							</td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
						</tr>
					</xsl:for-each>

					<tr>
						<td class="border">Total</td>
						<td class="border"></td>
						<td class="border"></td>
						<td class="border"></td>
						<td class="border"></td>
						<td class="border"></td>
						<td class="border"></td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Expenses", R[-<xsl:value-of select="$Rows" />]C)+SUMIF(R[-<xsl:value-of select="$Rows" />]C1:R[-1]C1, "Days", R[-<xsl:value-of select="$Rows" />]C[15])
						</td>
						<td class="border">
							=SUM(RC[-12]:RC[-1])
						</td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
					</tr>
				</xsl:for-each>
				<tr>
					<td></td>
				</tr>
			</xsl:for-each>
		</xsl:for-each>
	</xsl:template>

</xsl:stylesheet>
