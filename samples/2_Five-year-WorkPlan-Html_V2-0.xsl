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

	<xsl:template match="ConservationProject">
		<html>
			<head>
				<style>
					body
					{
					font-family: Arial;
					font-size: 10pt
					}

					table
					{
					padding: 0 2px 0 2px;
					}

					td
					{
					vertical-align: top;
					padding: 0 2pt 0 2pt;
					}

					td.borderheader
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					text-align: center;
					font-weight: bold;
					white-space: nowrap;
					};

					td.border
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					};

					.break { page-break-before: always; }
					.conservationtargetpicture { width: 5cm; }
					.coverpagedate { font-size: 14pt; text-align:center; }
					.coverpagepicture { text-align: center; }
					.coverpagetitle { font-size: 24pt; font-weight: bold; text-align: center}
					.instruction { font-style: italic; }
					.narrow {font-family: Arial Narrow}
					.nowrap {white-space: nowrap}
					.subheading { color: #4F81BD; font-weight: bold; }
					.toc {font-size: 13pt; font-weight: bold; color: #4F81BD; }
					.vision { font-style: italic; }
				</style>
			</head>
			<body>
				<xsl:apply-templates select="ProjectSummary" mode="titlepage" />
				<xsl:apply-templates select="ProjectSummary" mode="introduction"/>
				<xsl:apply-templates select="StrategyPool" mode="summary"/>
				<xsl:apply-templates select="ResultsChainPool" mode="resultschain"/>
				<xsl:apply-templates select="StrategyPool" mode="summaryplan"/>
			</body>
		</html>
	</xsl:template>

	<xsl:template match="ProjectSummary" mode="titlepage">
		<img src="http://localhost/Miradi/images/Logo.jpg" />
		<br />
		<br />
		<br />
		<br />
		<br />
		<br />
		<p class="coverpagetitle">
			<xsl:value-of select="ProjectSummaryProjectName"/>
		</p>
		<p class="coverpagetitle">
			Work Plan and Budget
		</p>
		<br />
		<br />
		<br />
		<br />
		<br />
		<br />
		<div class="coverpagepicture">
			<p class="instruction">Insert picture.</p>
		</div>
		<br />
		<br />
		<br />
		<br />
		<br />
		<br />
		<div class="coverpagedate">
			<p class="instruction">Insert date.</p>
		</div>
		<div class="break">
			<span class="toc">Table of Contents</span>
		</div>
		<p class="instruction">
			Insert table of contents (in Word: References,Table of Contents)
		</p>
	</xsl:template>

	<xsl:template match="ProjectSummary" mode="introduction">
		<h1 class="break">
			<xsl:value-of select="ProjectSummaryProjectName"/> Work Plan and Budget
		</h1>
		<p>
			This document outlines the work plan and budget for activities occurring at <xsl:value-of select="ProjectSummaryProjectName"/>. This Plan has been signed-off by Bush Heritage management during development of the Property Plan, and will be updated periodically as part of our Conservation Management Process which reflects our adaptive management philosophy.  More details are held in the Miradi database, and this should be referenced for the most up-to-date information.
		</p>
		<p class="instruction">
			Insert map showing location of targets (and threats?)
		</p>
	</xsl:template>

	<xsl:template match="StrategyPool" mode="summary">
		<h1 class="break">
			Workplan Summary
		</h1>
		<table>
			<tr>
				<td class="borderheader">
					Strategy
				</td>
				<td class="borderheader">
					Objective
				</td>
				<td class="borderheader">
				</td>
				<xsl:call-template name="for-workplansummaryheader">
					<xsl:with-param name="StartYear">
						<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanStartDate,1,4)"/>
					</xsl:with-param>
					<xsl:with-param name="StopYear">
						<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanEndDate,1,4) - 1"/>
					</xsl:with-param>
				</xsl:call-template>
				<td class="borderheader">
					Total
				</td>
			</tr>
			<xsl:for-each select="Strategy">
				<xsl:sort select="StrategyId" />
				<xsl:variable name="ObjectiveId" select="StrategyObjectiveIds/ObjectiveId" />
				<tr>
					<td class="border">
						<b>
							<xsl:value-of select="StrategyName"/>
						</b>
					</td>
					<td class="border">
						<table>
							<xsl:for-each select="/ConservationProject/ObjectivePool/Objective[@Id = $ObjectiveId]/ObjectiveName">
								<tr>
									<td>
										<xsl:value-of select="text()"/>
									</td>
								</tr>
							</xsl:for-each>
						</table>
					</td>
					<td class="border" align="right">
						<xsl:call-template name="ExpenseRowHeaders" />
					</td>
					<xsl:call-template name="for-workplansummaryrows">
						<xsl:with-param name="StartYear">
							<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanStartDate,1,4)"/>
						</xsl:with-param>
						<xsl:with-param name="StopYear">
							<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanEndDate,1,4) - 1"/>
						</xsl:with-param>
					</xsl:call-template>
					<td class="border" align="right">
						<xsl:variable name="StartYear" select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanStartDate,1,4)" />
						<xsl:variable name="StopYear" select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanEndDate,1,4) - 1" />
						<xsl:variable name="TaskId" select="StrategySortedActivityIds/ActivityId" />
						<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId" />
						<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId" />
						<xsl:variable name="NumberOfUnits" select="sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]/../../NumberOfUnits) + sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]/../../NumberOfUnits)" />
						<xsl:if test="$NumberOfUnits &gt; 0">
							<xsl:value-of select="$NumberOfUnits"/>
						</xsl:if>

						<xsl:variable name="AllUnits">
							<months>
								<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]">
									<unit_cost>
										<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
										<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
										<xsl:if test="not ($DailyRate)">
											<xsl:value-of select="0"/>
										</xsl:if>
										<xsl:if test="$DailyRate">
											<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
										</xsl:if>
									</unit_cost>
								</xsl:for-each>
							</months>
							<years>
								<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]">
									<unit_cost>
										<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
										<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
										<xsl:if test="not ($DailyRate)">
											<xsl:value-of select="0"/>
										</xsl:if>
										<xsl:if test="$DailyRate">
											<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
										</xsl:if>
									</unit_cost>
								</xsl:for-each>
							</years>
						</xsl:variable>
						<xsl:variable name="TotalUnitCost" select="msxsl:node-set($AllUnits)"/>
						<xsl:variable name="UnitsCost" select="sum($TotalUnitCost/months/unit_cost) + sum($TotalUnitCost/years/unit_cost)" />

						<br />
						<xsl:variable name="Expense" select="sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]/../../Expense) + sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]/../../Expense)" />
						<xsl:if test="$Expense &gt; 0">
							<xsl:value-of select="$Expense"/>							
						</xsl:if>
						<br />
						<xsl:variable name="Total" select="$UnitsCost + $Expense" />
						<xsl:if test="$Total &gt; 0 and $Total != 'NaN'">
							<xsl:value-of select="$Total"/>
						</xsl:if>
					</td>
				</tr>
			</xsl:for-each>
			<tr>
				<td class="border">
				</td>
				<td class="border" align="right">
					TOTAL
				</td>
				<td class="border" align="right">
					<xsl:call-template name="ExpenseRowHeaders" />
				</td>
				<xsl:call-template name="for-workplansummarytotals">
					<xsl:with-param name="StartYear">
						<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanStartDate,1,4)"/>
					</xsl:with-param>
					<xsl:with-param name="StopYear">
						<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanEndDate,1,4) - 1"/>
					</xsl:with-param>
				</xsl:call-template>
				<td class="border" align="right">
					<xsl:variable name="StartYear" select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanStartDate,1,4)" />
					<xsl:variable name="StopYear" select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanEndDate,1,4) - 1" />
					<xsl:variable name="TaskId" select="Strategy/StrategySortedActivityIds/ActivityId" />
					<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId" />
					<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId" />
					<xsl:variable name="TotalNumberOfUnits" select="sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]/../../NumberOfUnits) + sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]/../../NumberOfUnits)" />
					<xsl:if test="$TotalNumberOfUnits &gt; 0">
						<xsl:value-of select="$TotalNumberOfUnits"/>
					</xsl:if>

					<xsl:variable name="AllUnits">
						<months>
							<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]">
								<unit_cost>
									<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
									<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
									<xsl:if test="not ($DailyRate)">
										<xsl:value-of select="0"/>
									</xsl:if>
									<xsl:if test="$DailyRate">
										<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
									</xsl:if>
								</unit_cost>
							</xsl:for-each>
						</months>
						<years>
							<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]">
								<unit_cost>
									<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
									<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
									<xsl:if test="not ($DailyRate)">
										<xsl:value-of select="0"/>
									</xsl:if>
									<xsl:if test="$DailyRate">
										<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
									</xsl:if>
								</unit_cost>
							</xsl:for-each>
						</years>
					</xsl:variable>
					<xsl:variable name="TotalUnitCost" select="msxsl:node-set($AllUnits)"/>
					<xsl:variable name="TotalUnitsCost" select="sum($TotalUnitCost/months/unit_cost) + sum($TotalUnitCost/years/unit_cost)" />

					<br />
					<xsl:variable name="TotalExpense" select="sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]/../../Expense) + sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]/../../Expense)" />
					<xsl:if test="$TotalExpense &gt; 0">
						<xsl:value-of select="$TotalExpense"/>
					</xsl:if>
					<br />
					<xsl:variable name="GrandTotal" select="$TotalUnitsCost + $TotalExpense" />
					<xsl:if test="$GrandTotal &gt; 0 and $GrandTotal != 'NaN'">
						<xsl:value-of select="$GrandTotal"/>
					</xsl:if>
				</td>
			</tr>
		</table>
	</xsl:template>

	<xsl:template name="for-workplansummaryheader">
		<xsl:param name="StartYear" />
		<xsl:param name="StopYear" />
		<xsl:if test="$StartYear &lt;= $StopYear">
			<td class="borderheader">
				<xsl:value-of select="$StartYear - 2000" />-<xsl:value-of select="$StartYear - 2000 + 1" />
			</td>
		</xsl:if>
		<xsl:if test="$StartYear &lt;= $StopYear">
			<xsl:call-template name="for-workplansummaryheader">
				<xsl:with-param name="StartYear">
					<xsl:value-of select="$StartYear + 1"/>
				</xsl:with-param>
				<xsl:with-param name="StopYear">
					<xsl:value-of select="$StopYear"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template name="for-workplansummaryrows">
		<xsl:param name="StartYear" />
		<xsl:param name="StopYear" />
		<xsl:if test="$StartYear &lt;= $StopYear">
			<td class="border" align="right">
				<xsl:variable name="TaskId" select="StrategySortedActivityIds/ActivityId" />
				<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId" />
				<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId" />
				<xsl:variable name="NumberOfUnits" select="sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]/../../NumberOfUnits) + sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]/../../NumberOfUnits)" />
				<xsl:if test="$NumberOfUnits &gt; 0">
					<xsl:value-of select="$NumberOfUnits"/>
				</xsl:if>

				<xsl:variable name="AllUnits">
					<months>
						<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]">
							<unit_cost>
								<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
								<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
								<xsl:if test="not ($DailyRate)">
									<xsl:value-of select="0"/>
								</xsl:if>
								<xsl:if test="$DailyRate">
									<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
								</xsl:if>
							</unit_cost>
						</xsl:for-each>
					</months>
					<years>
						<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]">
							<unit_cost>
								<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
								<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
								<xsl:if test="not ($DailyRate)">
									<xsl:value-of select="0"/>
								</xsl:if>
								<xsl:if test="$DailyRate">
									<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
								</xsl:if>
							</unit_cost>
						</xsl:for-each>
					</years>
				</xsl:variable>
				<xsl:variable name="TotalUnitCost" select="msxsl:node-set($AllUnits)"/>
				<xsl:variable name="UnitsCost" select="sum($TotalUnitCost/months/unit_cost) + sum($TotalUnitCost/years/unit_cost)" />

				<br />
				<xsl:variable name="Expense" select="sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]/../../Expense) + sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]/../../Expense)" />
				<xsl:if test="$Expense &gt; 0">
					<xsl:value-of select="$Expense"/>
				</xsl:if>
				<br />
				<xsl:variable name="Total" select="$UnitsCost + $Expense" />
				<xsl:if test="$Total &gt; 0 and $Total != 'NaN'">
					<xsl:value-of select="$Total"/>
				</xsl:if>
			</td>
		</xsl:if>
		<xsl:if test="$StartYear &lt;= $StopYear">
			<xsl:call-template name="for-workplansummaryrows">
				<xsl:with-param name="StartYear">
					<xsl:value-of select="$StartYear + 1"/>
				</xsl:with-param>
				<xsl:with-param name="StopYear">
					<xsl:value-of select="$StopYear"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template name="for-workplansummarytotals">
		<xsl:param name="StartYear" />
		<xsl:param name="StopYear" />
		<xsl:if test="$StartYear &lt;= $StopYear">
			<td class="border" align="right">
				<xsl:variable name="TaskId" select="Strategy/StrategySortedActivityIds/ActivityId" />
				<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId" />
				<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId" />
				<xsl:variable name="NumberOfUnits" select="sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]/../../NumberOfUnits) + sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]/../../NumberOfUnits)" />
				<xsl:if test="$NumberOfUnits &gt; 0">
					<xsl:value-of select="$NumberOfUnits"/>
				</xsl:if>

				<xsl:variable name="AllUnits">
					<months>
						<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]">
							<unit_cost>
								<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
								<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
								<xsl:if test="not ($DailyRate)">
									<xsl:value-of select="0"/>
								</xsl:if>
								<xsl:if test="$DailyRate">
									<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
								</xsl:if>
							</unit_cost>
						</xsl:for-each>
					</months>
					<years>
						<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]">
							<unit_cost>
								<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
								<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
								<xsl:if test="not ($DailyRate)">
									<xsl:value-of select="0"/>
								</xsl:if>
								<xsl:if test="$DailyRate">
									<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
								</xsl:if>
							</unit_cost>
						</xsl:for-each>
					</years>
				</xsl:variable>
				<xsl:variable name="TotalUnitCost" select="msxsl:node-set($AllUnits)"/>
				<xsl:variable name="UnitsCost" select="sum($TotalUnitCost/months/unit_cost) + sum($TotalUnitCost/years/unit_cost)" />

				<br />
				<xsl:variable name="Expense" select="sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]/../../Expense) + sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]/../../Expense)" />
				<xsl:if test="$Expense &gt; 0">
					<xsl:value-of select="$Expense"/>
				</xsl:if>
				<br />
				<xsl:variable name="Total" select="$UnitsCost + $Expense" />
				<xsl:if test="$Total &gt; 0 and $Total != 'NaN'">
					<xsl:value-of select="$Total"/>
				</xsl:if>
			</td>
		</xsl:if>
		<xsl:if test="$StartYear &lt;= $StopYear">
			<xsl:call-template name="for-workplansummarytotals">
				<xsl:with-param name="StartYear">
					<xsl:value-of select="$StartYear + 1"/>
				</xsl:with-param>
				<xsl:with-param name="StopYear">
					<xsl:value-of select="$StopYear"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template name="ExpenseRowHeaders">
		Days<br />Expenses<br />Total
	</xsl:template>

	<xsl:template match="ResultsChainPool" mode="resultschain">
		<xsl:for-each select="ResultsChain">
			<xsl:sort select="ResultsChainId" />
			<h1 class="break">
				<xsl:value-of select="ResultsChainName"/>
			</h1>
			<xsl:variable name="DiagramFactorId" select="ResultsChainDiagramFactorIds/DiagramFactorId" />
			<xsl:variable name="StrategyId" select="/ConservationProject/DiagramFactorPool/DiagramFactor[@Id = $DiagramFactorId]/DiagramFactorWrappedFactorId/WrappedByDiagramFactorId/StrategyId" />
			<xsl:for-each select="/ConservationProject/StrategyPool/Strategy[@Id = $StrategyId]">
				<h2>
					<xsl:value-of select="StrategyName"/>
				</h2>
				<p>
					<xsl:value-of select="StrategyDetails"/>
				</p>
				<table>
					<tr>
						<td>
							<b>Objectives:</b>
						</td>
						<td>
							<table>
								<xsl:for-each select="StrategyObjectiveIds/ObjectiveId">
									<xsl:call-template name="objectiverow">
										<xsl:with-param name="ObjectiveId">
											<xsl:value-of select="text()"/>
										</xsl:with-param>
									</xsl:call-template>
								</xsl:for-each>
							</table>
						</td>
					</tr>
				</table>
				<table>
					<tr>
						<td class="borderheader">
							Activity
						</td>
						<td class="borderheader">
							Description
						</td>
						<td class="borderheader">
						</td>
						<xsl:call-template name="for-workplansummaryheader">
							<xsl:with-param name="StartYear">
								<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanStartDate,1,4)"/>
							</xsl:with-param>
							<xsl:with-param name="StopYear">
								<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanEndDate,1,4) - 1"/>
							</xsl:with-param>
						</xsl:call-template>
						<td class="borderheader">
							Total
						</td>
					</tr>
					<tr>
						<td class="border">
							<b>
								<xsl:value-of select="StrategyName"/>
							</b>
						</td>
						<td class="border" align="right">
							<b>
								Aggregated Effort, Expenses and Total Cost for this Strategy
							</b>
						</td>
						<td class="border" align="right">
							<b>
								<xsl:call-template name="ExpenseRowHeaders" />
							</b>
						</td>
						<xsl:call-template name="for-strategysummaryrows">
							<xsl:with-param name="StartYear">
								<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanStartDate,1,4)"/>
							</xsl:with-param>
							<xsl:with-param name="StopYear">
								<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanEndDate,1,4) - 1"/>
							</xsl:with-param>
						</xsl:call-template>
						<td class="border" align="right">
							<b>
								<xsl:variable name="StartYear" select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanStartDate,1,4)" />
								<xsl:variable name="StopYear" select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanEndDate,1,4) - 1" />
								<xsl:variable name="TaskId" select="StrategySortedActivityIds/ActivityId" />
								<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId" />
								<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId" />
								<xsl:variable name="TotalNumberOfUnits" select="sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]/../../NumberOfUnits) + sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]/../../NumberOfUnits)" />
								<xsl:if test="$TotalNumberOfUnits &gt; 0">
									<xsl:value-of select="$TotalNumberOfUnits"/>
								</xsl:if>

								<xsl:variable name="AllUnits">
									<months>
										<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year &gt;= $StartYear) and (@Year &lt;= $StopYear+1)]">
											<unit_cost>
												<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
												<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
												<xsl:if test="not ($DailyRate)">
													<xsl:value-of select="0"/>
												</xsl:if>
												<xsl:if test="$DailyRate">
													<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
												</xsl:if>
											</unit_cost>
										</xsl:for-each>
									</months>
									<years>
										<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear &gt;= $StartYear) and (@StartYear &lt;= $StopYear+1)]">
											<unit_cost>
												<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
												<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
												<xsl:if test="not ($DailyRate)">
													<xsl:value-of select="0"/>
												</xsl:if>
												<xsl:if test="$DailyRate">
													<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
												</xsl:if>
											</unit_cost>
										</xsl:for-each>
									</years>
								</xsl:variable>
								<xsl:variable name="TotalUnitCost" select="msxsl:node-set($AllUnits)"/>
								<xsl:variable name="TotalUnitsCost" select="sum($TotalUnitCost/months/unit_cost) + sum($TotalUnitCost/years/unit_cost)" />

								<br />
								<xsl:variable name="TotalExpense" select="sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]/../../Expense) + sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]/../../Expense)" />
								<xsl:if test="$TotalExpense &gt; 0">
									<xsl:value-of select="$TotalExpense"/>
								</xsl:if>
								<br />
								<xsl:variable name="GrandTotal" select="$TotalUnitsCost + $TotalExpense" />
								<xsl:if test="$GrandTotal &gt; 0">
									<xsl:value-of select="$GrandTotal"/>
								</xsl:if>
							</b>
						</td>
					</tr>
					<xsl:for-each select="StrategySortedActivityIds/ActivityId">
						<xsl:variable name="TaskId">
							<xsl:value-of select="text()"/>
						</xsl:variable>
						<tr>
							<td class="border">
								<xsl:value-of select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskName"/>
							</td>
							<td class="border">
								<xsl:value-of select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskDetails"/>
							</td>
							<td class="border" align="right">
								<xsl:call-template name="ExpenseRowHeaders" />
							</td>
							<xsl:call-template name="for-activity">
								<xsl:with-param name="StartYear">
									<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanStartDate,1,4)"/>
								</xsl:with-param>
								<xsl:with-param name="StopYear">
									<xsl:value-of select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanEndDate,1,4) - 1"/>
								</xsl:with-param>
								<xsl:with-param name="TaskId">
									<xsl:value-of select="$TaskId"/>
								</xsl:with-param>
							</xsl:call-template>
							<td class="border" align="right">
								<xsl:variable name="StartYear" select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanStartDate,1,4)" />
								<xsl:variable name="StopYear" select="substring(/ConservationProject/ProjectSummaryPlanning/ProjectSummaryPlanningWorkPlanEndDate,1,4) - 1" />
								<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId" />
								<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId" />
								<xsl:variable name="NumberOfUnits" select="sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]/../../NumberOfUnits) + sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]/../../NumberOfUnits)"/>
								<xsl:if test="$NumberOfUnits &gt; 0">
									<xsl:value-of select="$NumberOfUnits"/>
								</xsl:if>

								<xsl:variable name="AllUnits">
									<months>
										<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]">
											<unit_cost>
												<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
												<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
												<xsl:if test="not ($DailyRate)">
													<xsl:value-of select="0"/>
												</xsl:if>
												<xsl:if test="$DailyRate">
													<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
												</xsl:if>
											</unit_cost>
										</xsl:for-each>
									</months>
									<years>
										<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]">
											<unit_cost>
												<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
												<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
												<xsl:if test="not ($DailyRate)">
													<xsl:value-of select="0"/>
												</xsl:if>
												<xsl:if test="$DailyRate">
													<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
												</xsl:if>
											</unit_cost>
										</xsl:for-each>
									</years>
								</xsl:variable>
								<xsl:variable name="TotalUnitCost" select="msxsl:node-set($AllUnits)"/>
								<xsl:variable name="UnitsCost" select="sum($TotalUnitCost/months/unit_cost) + sum($TotalUnitCost/years/unit_cost)" />

								<br />
								<xsl:variable name="Expense" select="sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year &gt;= $StartYear and @Year &lt;= $StopYear+1)]/../../Expense) + sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear &gt;= $StartYear and @StartYear &lt;= $StopYear+1)]/../../Expense)"/>
								<xsl:if test="$Expense &gt; 0">
									<xsl:value-of select="$Expense"/>
								</xsl:if>
								<br />
								<xsl:variable name="Total" select="$UnitsCost + $Expense"/>
								<xsl:if test="$Total &gt; 0 and $Total != 'NaN'">
									<xsl:value-of select="$Total"/>
								</xsl:if>
							</td>
						</tr>
					</xsl:for-each>
				</table>
				<h4>Monitoring</h4>
				<xsl:call-template name="indicators">
					<xsl:with-param name="StrategyId">
						<xsl:value-of select="@Id"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:for-each>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="objectiverow">
		<xsl:param name="ObjectiveId" />
		<tr>
			<td>
				<xsl:value-of select="/ConservationProject/ObjectivePool/Objective[@Id=$ObjectiveId]/ObjectiveName" />
			</td>
			<td>
				<xsl:value-of select="/ConservationProject/ObjectivePool/Objective[@Id=$ObjectiveId]/ObjectiveDetails" />
			</td>
		</tr>
	</xsl:template>

	<xsl:template name="resultschainname">
		<xsl:param name="DiagramFactorId" />
		<xsl:for-each select="/ConservationProject/ResultsChainPool/ResultsChain">
			<xsl:if test="ResultsChainDiagramFactorIds/DiagramFactorId = $DiagramFactorId">
				<xsl:value-of select="ResultsChainName"/>
				<br />
			</xsl:if>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="for-strategysummaryrows">
		<xsl:param name="StartYear" />
		<xsl:param name="StopYear" />
		<xsl:if test="$StartYear &lt;= $StopYear">
			<td class="border" align="right">
				<b>
					<xsl:variable name="TaskId" select="StrategySortedActivityIds/ActivityId" />
					<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId" />
					<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId" />
					<xsl:variable name="NumberOfUnits" select="sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]/../../NumberOfUnits) + sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]/../../NumberOfUnits)" />
					<xsl:if test="$NumberOfUnits &gt; 0">
						<xsl:value-of select="$NumberOfUnits"/>
					</xsl:if>

					<xsl:variable name="AllUnits">
						<months>
							<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]">
								<unit_cost>
									<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
									<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
									<xsl:if test="not ($DailyRate)">
										<xsl:value-of select="0"/>
									</xsl:if>
									<xsl:if test="$DailyRate">
										<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
									</xsl:if>
								</unit_cost>
							</xsl:for-each>
						</months>
						<years>
							<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]">
								<unit_cost>
									<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
									<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
									<xsl:if test="not ($DailyRate)">
										<xsl:value-of select="0"/>
									</xsl:if>
									<xsl:if test="$DailyRate">
										<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
									</xsl:if>
								</unit_cost>
							</xsl:for-each>
						</years>
					</xsl:variable>
					<xsl:variable name="TotalUnitCost" select="msxsl:node-set($AllUnits)"/>
					<xsl:variable name="UnitsCost" select="sum($TotalUnitCost/months/unit_cost) + sum($TotalUnitCost/years/unit_cost)" />

					<br />
					<xsl:variable name="Expense" select="sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]/../../Expense) + sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]/../../Expense)" />
					<xsl:if test="$Expense &gt; 0">
						<xsl:value-of select="$Expense"/>
					</xsl:if>
					<br />
					<xsl:variable name="Total" select="$UnitsCost + $Expense" />
					<xsl:if test="$Total &gt; 0 and $Total != 'NaN'">
						<xsl:value-of select="$Total"/>
					</xsl:if>
				</b>
			</td>
		</xsl:if>
		<xsl:if test="$StartYear &lt;= $StopYear">
			<xsl:call-template name="for-strategysummaryrows">
				<xsl:with-param name="StartYear">
					<xsl:value-of select="$StartYear + 1"/>
				</xsl:with-param>
				<xsl:with-param name="StopYear">
					<xsl:value-of select="$StopYear"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template name="for-activity">
		<xsl:param name="StartYear" />
		<xsl:param name="StopYear" />
		<xsl:param name="TaskId" />
		<xsl:if test="$StartYear &lt;= $StopYear">
			<td class="border" align="right">
				<xsl:variable name="ExpenseAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskExpenseIds/ExpenseAssignmentId" />
				<xsl:variable name="ResourceAssignmentId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskAssignmentIds/ResourceAssignmentId" />
				<xsl:variable name="NumberOfUnits" select="sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]/../../NumberOfUnits) + sum(/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]/../../NumberOfUnits)"/>
				<xsl:if test="$NumberOfUnits &gt; 0">
					<xsl:value-of select="$NumberOfUnits"/>
				</xsl:if>

				<xsl:variable name="AllUnits">
					<months>
						<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]">
							<unit_cost>
								<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
								<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
								<xsl:if test="not ($DailyRate)">
									<xsl:value-of select="0"/>
								</xsl:if>
								<xsl:if test="$DailyRate">
									<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
								</xsl:if>
							</unit_cost>
						</xsl:for-each>
					</months>
					<years>
						<xsl:for-each select="/ConservationProject/ResourceAssignmentPool/ResourceAssignment[@Id = $ResourceAssignmentId]/ResourceAssignmentDetails/DateUnitWorkUnits/WorkUnitsDateUnit/WorkUnitsYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]">
							<unit_cost>
								<xsl:variable name="ResourceId" select="../../../../ResourceAssignmentResourceId/ResourceId" />
								<xsl:variable name="DailyRate" select="/ConservationProject/ProjectResourcePool/ProjectResource[@Id = $ResourceId]/ProjectResourceDailyRate" />
								<xsl:if test="not ($DailyRate)">
									<xsl:value-of select="0"/>
								</xsl:if>
								<xsl:if test="$DailyRate">
									<xsl:value-of select="../../NumberOfUnits * $DailyRate"/>
								</xsl:if>
							</unit_cost>
						</xsl:for-each>
					</years>
				</xsl:variable>
				<xsl:variable name="TotalUnitCost" select="msxsl:node-set($AllUnits)"/>
				<xsl:variable name="UnitsCost" select="sum($TotalUnitCost/months/unit_cost) + sum($TotalUnitCost/years/unit_cost)" />
				<br />
				<xsl:variable name="Expense" select="sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesMonth[(@Year = $StartYear and @Month &gt;= 7) or (@Year = $StartYear+1 and @Month &lt;= 6)]/../../Expense) + sum(/ConservationProject/ExpenseAssignmentPool/ExpenseAssignment[@Id = $ExpenseAssignmentId]/ExpenseAssignmentDetails/DateUnitExpense/ExpensesDateUnit/ExpensesYear[(@StartYear = $StartYear and @StartMonth &gt;= 7) or (@StartYear = $StartYear+1 and @StartMonth &lt;= 6)]/../../Expense)"/>
				<xsl:if test="$Expense &gt; 0">
					<xsl:value-of select="$Expense"/>
				</xsl:if>
				<br />
				<xsl:variable name="Total" select="$UnitsCost + $Expense" />
				<xsl:if test="$Total &gt; 0 and $Total != 'NaN'">
					<xsl:value-of select="$Total"/>
				</xsl:if>
			</td>
		</xsl:if>
		<xsl:if test="$StartYear &lt;= $StopYear">
			<xsl:call-template name="for-activity">
				<xsl:with-param name="StartYear">
					<xsl:value-of select="$StartYear + 1"/>
				</xsl:with-param>
				<xsl:with-param name="StopYear">
					<xsl:value-of select="$StopYear"/>
				</xsl:with-param>
				<xsl:with-param name="TaskId">
					<xsl:value-of select="$TaskId"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template name="indicators">
		<xsl:param name="StrategyId" />
		<table>
			<tr>
				<td class="borderheader">
					Indicator
				</td>
				<td class="borderheader">
					Description
				</td>
				<td class="borderheader">
					Method
				</td>
				<td class="borderheader">
					Current Measure
				</td>
				<td class="borderheader">
					Desired Measure
				</td>
			</tr>
			<xsl:call-template name="indicatorrows">
				<xsl:with-param name="IndicatorId" select="/ConservationProject/StrategyPool/Strategy[@Id = $StrategyId]/StrategyIndicatorIds/IndicatorId" />
			</xsl:call-template>
			<xsl:call-template name="intermediaterows">
				<xsl:with-param name="StrategyId" select="$StrategyId" />
			</xsl:call-template>
		</table>
	</xsl:template>

	<xsl:template name="indicatorrows">
		<xsl:param name="IndicatorId" />
		<xsl:for-each select="/ConservationProject/IndicatorPool/Indicator[@Id = $IndicatorId]">
			<tr>
				<td class="border">
					<xsl:value-of select="IndicatorName"/>
				</td>
				<td class="border">
					<xsl:value-of select="IndicatorDetails"/>
				</td>
				<td class="border">
					<xsl:call-template name="methodslist">
						<xsl:with-param name="IndicatorId" select="$IndicatorId" />
					</xsl:call-template>
				</td>
				<td class="border">
					<xsl:call-template name="indicatorslist">
						<xsl:with-param name="IndicatorId" select="$IndicatorId" />
					</xsl:call-template>
				</td>
				<td class="border">
					<table>
						<tr>
							<td class="nowrap">
								<xsl:value-of select="IndicatorFutureStatusDate"/>
							</td>
							<td class="nowrap">
								<xsl:value-of select="IndicatorFutureStatusSummary"/>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="intermediaterows">
		<xsl:param name="StrategyId" />
		<xsl:for-each select="/ConservationProject/DiagramFactorPool/DiagramFactor">
			<xsl:if test="DiagramFactorWrappedFactorId/WrappedByDiagramFactorId/StrategyId = $StrategyId">
				<xsl:variable name="DiagramLinkStrategyId" select="@Id" />
				<xsl:for-each select="/ConservationProject/DiagramLinkPool/DiagramLink">
					<xsl:if test="DiagramLinkFromDiagramFactorId/LinkableFactorId/StrategyId = $DiagramLinkStrategyId">
						<xsl:variable name="DiagramFactorIntermediateResultId" select="DiagramLinkToDiagramFactorId/LinkableFactorId/IntermediateResultId" />
						<xsl:variable name="IntermediateResultId" select="/ConservationProject/DiagramFactorPool/DiagramFactor[@Id=$DiagramFactorIntermediateResultId]/DiagramFactorWrappedFactorId/WrappedByDiagramFactorId/IntermediateResultId" />
						<xsl:variable name="IndicatorId" select="/ConservationProject/IntermediateResultPool/IntermediateResult[@Id = $IntermediateResultId]/IntermediateResultIndicatorIds/IndicatorId" />
						<xsl:call-template name="indicatorrows">
							<xsl:with-param name="IndicatorId" select="$IndicatorId" />
						</xsl:call-template>
					</xsl:if>
				</xsl:for-each>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="methodslist">
		<xsl:param name="IndicatorId" />
		<table>
			<xsl:for-each select="/ConservationProject/IndicatorPool/Indicator[@Id = $IndicatorId]/IndicatorSortedMethodIds/MethodId">
				<tr>
					<td>
						<xsl:variable name="MethodId" select="text()" />
						<xsl:value-of select="/ConservationProject/TaskPool/Task[@Id = $MethodId]/TaskName"/>
					</td>
				</tr>
			</xsl:for-each>
		</table>
	</xsl:template>

	<xsl:template name="indicatorslist">
		<xsl:param name="IndicatorId" />
		<table>
			<xsl:for-each select="/ConservationProject/IndicatorPool/Indicator[@Id = $IndicatorId]/IndicatorMeasurementIds/MeasurementId">
				<xsl:variable name="MeasurementId" select="text()" />
				<tr>
					<td class="nowrap">
						<xsl:value-of select="/ConservationProject/MeasurementPool/Measurement[@Id = $MeasurementId]/MeasurementDate"/>
					</td>
					<td class="nowrap">
						<xsl:value-of select="/ConservationProject/MeasurementPool/Measurement[@Id = $MeasurementId]/MeasurementMeasurementValue"/>
					</td>
				</tr>
			</xsl:for-each>
		</table>
	</xsl:template>

	<xsl:template match="StrategyPool" mode="summaryplan">
		<h1 class="break">
			Summary Monitoring Plan
		</h1>
		<p>
			The Table below expands on the "Measuring Our Progress" table in the Management Plan document, and summarises the progress we are making towards restoring and protecting the key values of the property. The first column documents each Conservation Target, then the Key Ecological Attributes of the Target, the Indicators that we will monitor to assess the health of those Attributes, the most recent measurement for the Indicator, and the Desired Future Value of the indicator. Other columns then show the current Status of the Indicator, four states of the indicator ranging from Poor to Very Good, and finally the source of measurements for the indicator.
		</p>
		<p class="instruction">Manually insert Target Viability Table from Miradi Reports.</p>
	</xsl:template>
	
</xsl:stylesheet>
