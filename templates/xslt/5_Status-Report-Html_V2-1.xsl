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

					td.borderheaderactivity
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					text-align: center;
					font-weight: bold;
					white-space: nowrap;
					width: 200pt;
					};

					td.borderheaderdescription
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					text-align: center;
					font-weight: bold;
					white-space: nowrap;
					width: 400pt;
					};

					td.border
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					};

					td.bordernowrap
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					white-space: nowrap;
					};

					td.borderprogressstatusplanned
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					background-color: #E0E0E0;
					};

					td.borderprogressstatusmajorissues
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					background-color: #E90A03;
					};

					td.borderprogressstatusminorissues
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					background-color: #FFED1E;
					};

					td.borderprogressstatusontrack
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					background-color: #76FF00;
					};

					td.borderprogressstatuscompleted
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					background-color: #039F00;
					};

					td.borderprogressstatusabandoned
					{
					border-collapse: collapse;
					border-style: solid;
					border-width: 1px;
					border-color: black;
					background-color: #808080;
					};

					.break
					{
					page-break-before: always;
					}

					.center
					{
					text-align: center;
					}

					.conservationtargetpicture
					{
					width: 5cm;
					}

					.coverpagedate
					{
					font-size: 14pt;
					text-align:center;
					}

					.coverpagetitle
					{
					color: #234170;
					font-size: 24pt;
					font-weight: bold;
					text-align: center
					}

					.instruction
					{
					font-style: italic;
					}

					.narrow
					{
					font-family: Arial Narrow
					}

					.nowrap
					{
					white-space: nowrap
					}

					.subheading
					{
					color: #4F81BD;
					font-weight: bold;
					}

					.toc
					{
					font-size: 13pt;
					font-weight:
					bold; color: #4F81BD;
					}

					.vision
					{
					font-style: italic;
					}

				</style>
			</head>
			<body>
				<xsl:apply-templates select="ProjectSummary" mode="titlepage" />
				<xsl:apply-templates select="ResultsChainPool" mode="summary"/>
				<xsl:apply-templates select="ResultsChainPool" mode="resultschain"/>
			</body>
		</html>
	</xsl:template>

	<xsl:template match="ProjectSummary" mode="titlepage">
		<p class="center"><img src="http://localhost/Miradi/images/Logo.jpg" /></p>
		<br />
		<br />
		<br />
		<br />
		<br />
		<br />
		<p class="coverpagetitle">
			<xsl:value-of select="ProjectSummaryProjectName"/> Status Report
		</p>
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
			<span class="toc">Summary - Status of each Strategy</span>
		</div>
	</xsl:template>

	<xsl:template match="ResultsChainPool" mode="summary">
		<table>
			<tr>
				<td class="borderheader">
					Strategy
				</td>
				<td class="borderheader">
					Objective
				</td>
				<td class="borderheader">
					Activity Code
				</td>
				<td class="borderheader">
					Date
				</td>
				<td class="borderheader">
					Status
				</td>
				<td class="borderheader">
					Details
				</td>
			</tr>
			<xsl:for-each select="ResultsChain">
				<xsl:sort select="ResultsChainId" />
				<xsl:variable name="DiagramFactorId" select="ResultsChainDiagramFactorIds/DiagramFactorId" />
				<xsl:variable name="StrategyId" select="/ConservationProject/DiagramFactorPool/DiagramFactor[@Id = $DiagramFactorId]/DiagramFactorWrappedFactorId/WrappedByDiagramFactorId/StrategyId" />
				<xsl:for-each select="/ConservationProject/StrategyPool/Strategy[@Id = $StrategyId]">
					<xsl:sort select="StrategyId" />
					<xsl:variable name="ObjectiveId" select="StrategyObjectiveIds/ObjectiveId" />
					<xsl:variable name="ProgressReportId" select="StrategyProgressReportIds/ProgressReportId" />
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
										<td class="border">
											<xsl:value-of select="text()"/>
										</td>
									</tr>
								</xsl:for-each>
							</table>
						</td>

						<td class="border">
							<xsl:value-of select="StrategyId"/>
						</td>

						<xsl:if test="count(/ConservationProject/ProgressReportPool/ProgressReport[@Id = $ProgressReportId]) = 0">
							<td class="bordernowrap">
							</td>
							<td class="border">
							</td>
							<td class="border">
							</td>
						</xsl:if>
						<xsl:for-each select="/ConservationProject/ProgressReportPool/ProgressReport[@Id = $ProgressReportId]">
							<xsl:sort select="ProgressReportProgressDate" />
							<xsl:if test="position() = last()">
								<td class="bordernowrap">
									<xsl:value-of select="ProgressReportProgressDate" />
								</td>
								<xsl:choose>
									<xsl:when test="ProgressReportProgressStatus = 'Planned'">
										<td class="borderprogressstatusplanned">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:when test="ProgressReportProgressStatus = 'MajorIssues'">
										<td class="borderprogressstatusmajorissues">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:when test="ProgressReportProgressStatus = 'MinorIssues'">
										<td class="borderprogressstatusminorissues">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:when test="ProgressReportProgressStatus = 'OnTrack'">
										<td class="borderprogressstatusontrack">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:when test="ProgressReportProgressStatus = 'Completed'">
										<td class="borderprogressstatuscompleted">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:when test="ProgressReportProgressStatus = 'Abandoned'">
										<td class="borderprogressstatusabandoned">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:otherwise>
										<td class="border">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:otherwise>
								</xsl:choose>
								<td class="border">
									<xsl:value-of select="ProgressReportDetails" />
								</td>
							</xsl:if>
						</xsl:for-each>

					</tr>
				</xsl:for-each>
			</xsl:for-each>
		</table>
	</xsl:template>

	<xsl:template match="ResultsChainPool" mode="resultschain">
		<xsl:for-each select="ResultsChain">
			<xsl:sort select="ResultsChainId" />
			<h1 class="break">
				Results Chain: <xsl:value-of select="ResultsChainName"/>
			</h1>
			<xsl:variable name="DiagramFactorId" select="ResultsChainDiagramFactorIds/DiagramFactorId" />
			<xsl:variable name="StrategyId" select="/ConservationProject/DiagramFactorPool/DiagramFactor[@Id = $DiagramFactorId]/DiagramFactorWrappedFactorId/WrappedByDiagramFactorId/StrategyId" />
			<xsl:for-each select="/ConservationProject/StrategyPool/Strategy[@Id = $StrategyId]">
				<xsl:sort select="StrategyId" />
				<xsl:variable name="ProgressReportId" select="StrategyProgressReportIds/ProgressReportId" />
				<h2>
					<xsl:value-of select="StrategyName"/>
				</h2>
				<table>
								<xsl:for-each select="StrategyObjectiveIds/ObjectiveId">
									<xsl:call-template name="objectiverow">
										<xsl:with-param name="ObjectiveId">
											<xsl:value-of select="text()"/>
										</xsl:with-param>
									</xsl:call-template>
								</xsl:for-each>
					<tr>
						<td class="borderheaderactivity">
							Activity
						</td>
						<td class="borderheaderdescription">
							Description
						</td>
						<td class="borderheader">
							Activity Code
						</td>
						<td class="borderheader">
							Date
						</td>
						<td class="borderheader">
							Status
						</td>
						<td class="borderheader">
							Details
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
								Overall Status
							</b>
						</td>
						<td class="border" align="right">
							<xsl:value-of select="StrategyId"/>
						</td>

						<xsl:for-each select="/ConservationProject/ProgressReportPool/ProgressReport[@Id = $ProgressReportId]">
							<xsl:sort select="ProgressReportProgressDate" />
							<xsl:if test="position() = last()">
								<td class="bordernowrap">
									<xsl:value-of select="ProgressReportProgressDate" />
								</td>
								<xsl:choose>
									<xsl:when test="ProgressReportProgressStatus = 'Planned'">
										<td class="borderprogressstatusplanned">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:when test="ProgressReportProgressStatus = 'MajorIssues'">
										<td class="borderprogressstatusmajorissues">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:when test="ProgressReportProgressStatus = 'MinorIssues'">
										<td class="borderprogressstatusminorissues">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:when test="ProgressReportProgressStatus = 'OnTrack'">
										<td class="borderprogressstatusontrack">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:when test="ProgressReportProgressStatus = 'Completed'">
										<td class="borderprogressstatuscompleted">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:when test="ProgressReportProgressStatus = 'Abandoned'">
										<td class="borderprogressstatusabandoned">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:when>
									<xsl:otherwise>
										<td class="border">
											<xsl:value-of select="ProgressReportProgressStatus" />
										</td>
									</xsl:otherwise>
								</xsl:choose>
								<td class="border">
									<xsl:value-of select="ProgressReportDetails" />
								</td>
							</xsl:if>
						</xsl:for-each>
						
					</tr>
					<xsl:for-each select="StrategySortedActivityIds/ActivityId">
						<xsl:variable name="TaskId">
							<xsl:value-of select="text()"/>
						</xsl:variable>
						<xsl:variable name="TaskProgressReportId" select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskProgressReportIds/ProgressReportId" />
						<xsl:if test="count(/ConservationProject/ProgressReportPool/ProgressReport[@Id = $TaskProgressReportId]) &gt; 0">
							<tr>
								<td class="border">
									<xsl:value-of select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskName"/>
								</td>
								<td class="border">
									<xsl:value-of select="/ConservationProject/TaskPool/Task[@Id = $TaskId]/TaskDetails"/>
								</td>
								<td class="border">
								</td>
								<xsl:for-each select="/ConservationProject/ProgressReportPool/ProgressReport[@Id = $TaskProgressReportId]">
									<xsl:sort select="ProgressReportProgressDate" />
									<xsl:if test="position() = last()">
										<td class="bordernowrap">
											<xsl:value-of select="ProgressReportProgressDate" />
										</td>
										<xsl:choose>
											<xsl:when test="ProgressReportProgressStatus = 'Planned'">
												<td class="borderprogressstatusplanned">
													<xsl:value-of select="ProgressReportProgressStatus" />
												</td>
											</xsl:when>
											<xsl:when test="ProgressReportProgressStatus = 'MajorIssues'">
												<td class="borderprogressstatusmajorissues">
													<xsl:value-of select="ProgressReportProgressStatus" />
												</td>
											</xsl:when>
											<xsl:when test="ProgressReportProgressStatus = 'MinorIssues'">
												<td class="borderprogressstatusminorissues">
													<xsl:value-of select="ProgressReportProgressStatus" />
												</td>
											</xsl:when>
											<xsl:when test="ProgressReportProgressStatus = 'OnTrack'">
												<td class="borderprogressstatusontrack">
													<xsl:value-of select="ProgressReportProgressStatus" />
												</td>
											</xsl:when>
											<xsl:when test="ProgressReportProgressStatus = 'Completed'">
												<td class="borderprogressstatuscompleted">
													<xsl:value-of select="ProgressReportProgressStatus" />
												</td>
											</xsl:when>
											<xsl:when test="ProgressReportProgressStatus = 'Abandoned'">
												<td class="borderprogressstatusabandoned">
													<xsl:value-of select="ProgressReportProgressStatus" />
												</td>
											</xsl:when>
											<xsl:otherwise>
												<td class="border">
													<xsl:value-of select="ProgressReportProgressStatus" />
												</td>
											</xsl:otherwise>
										</xsl:choose>
										<td class="border">
											<xsl:value-of select="ProgressReportDetails" />
										</td>
									</xsl:if>
								</xsl:for-each>
							</tr>
						</xsl:if>
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
				<b>Objective:</b>
			</td>
			<td colspan="5">
				<xsl:value-of select="/ConservationProject/ObjectivePool/Objective[@Id=$ObjectiveId]/ObjectiveName" />
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

</xsl:stylesheet>
