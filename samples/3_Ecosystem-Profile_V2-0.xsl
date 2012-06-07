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

	<xsl:template match="/">
    <xsl:processing-instruction name="mso-application">
			progid="Word.Document"
		</xsl:processing-instruction>
		<w:wordDocument
			xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" 
			xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint" 
			xmlns:o="urn:schemas-microsoft-com:office:office" 
			w:macrosPresent="no" 
			w:embeddedObjPresent="no" 
			w:ocxPresent="no" 
			xml:space="default">
			<w:fonts>
				<w:defaultFonts
					w:ascii="Times New Roman"
					w:fareast="Times New Roman"
					w:h-ansi="Times New Roman"
					w:cs="Times New Roman" />
			</w:fonts>
			<w:styles>
				<w:versionOfBuiltInStylenames w:val="4" />
				<w:latentStyles w:defLockedState="off" w:latentStyleCount="156" />
				<w:style w:type="paragraph" w:default="on" w:styleId="Normal">
					<w:name w:val="Normal" />
					<w:rPr>
						<wx:font wx:val="Times New Roman" />
						<w:sz w:val="24" />
						<w:sz-cs w:val="24" />
						<w:lang w:val="EN-US" w:fareast="EN-US" w:bidi="AR-SA" />
					</w:rPr>
				</w:style>
				<w:style w:type="paragraph" w:styleId="Heading1">
					<w:name w:val="heading 1" />
					<wx:uiName wx:val="Heading 1" />
					<w:basedOn w:val="Normal" />
					<w:next w:val="Normal" />
					<w:pPr>
						<w:pStyle w:val="Heading1" />
						<w:keepNext />
						<w:spacing w:before="240" w:after="60" />
						<w:outlineLvl w:val="0" />
					</w:pPr>
					<w:rPr>
						<w:rFonts w:ascii="Arial" w:h-ansi="Arial" w:cs="Arial" />
						<wx:font wx:val="Arial" />
						<w:b />
						<w:b-cs />
						<w:kern w:val="32" />
						<w:sz w:val="32" />
						<w:sz-cs w:val="32" />
					</w:rPr>
				</w:style>
				<w:style w:type="paragraph" w:styleId="Heading2">
					<w:name w:val="heading 2" />
					<wx:uiName wx:val="Heading 2" />
					<w:basedOn w:val="Normal" />
					<w:next w:val="Normal" />
					<w:pPr>
						<w:pStyle w:val="Heading2" />
						<w:keepNext />
						<w:spacing w:before="240" w:after="60" />
						<w:outlineLvl w:val="0" />
					</w:pPr>
					<w:rPr>
						<w:rFonts w:ascii="Arial" w:h-ansi="Arial" w:cs="Arial" />
						<wx:font wx:val="Arial" />
						<w:b />
						<w:b-cs />
						<w:kern w:val="24" />
						<w:sz w:val="24" />
						<w:sz-cs w:val="24" />
					</w:rPr>
				</w:style>
				<w:style w:type="character" w:default="on" w:styleId="DefaultParagraphFont">
					<w:name w:val="Default Paragraph Font" />
					<w:semiHidden />
				</w:style>
				<w:style w:type="table" w:default="on" w:styleId="TableNormal">
					<w:name w:val="Normal Table" />
					<wx:uiName wx:val="Table Normal" />
					<w:semiHidden />
					<w:rPr>
						<wx:font wx:val="Times New Roman" />
					</w:rPr>
					<w:tblPr>
						<w:tblInd w:w="0" w:type="dxa" />
						<w:tblCellMar>
							<w:top w:w="0" w:type="dxa" />
							<w:left w:w="108" w:type="dxa" />
							<w:bottom w:w="0" w:type="dxa" />
							<w:right w:w="108" w:type="dxa" />
						</w:tblCellMar>
					</w:tblPr>
				</w:style>
				<w:style w:type="list" w:default="on" w:styleId="NoList">
					<w:name w:val="No List" />
					<w:semiHidden />
				</w:style>
			</w:styles>
			<w:docPr>
				<w:view w:val="print" />
				<w:zoom w:percent="100" />
				<w:doNotEmbedSystemFonts />
				<w:proofState w:spelling="clean" w:grammar="clean" />
				<w:attachedTemplate w:val="" />
				<w:defaultTabStop w:val="720" />
				<w:punctuationKerning />
				<w:characterSpacingControl w:val="DontCompress" />
				<w:optimizeForBrowser />
				<w:validateAgainstSchema />
				<w:saveInvalidXML w:val="off" />
				<w:ignoreMixedContent w:val="off" />
				<w:alwaysShowPlaceholderText w:val="off" />
				<w:compat>
					<w:breakWrappedTables />
					<w:snapToGridInCell />
					<w:wrapTextWithPunct />
					<w:useAsianBreakRules />
					<w:dontGrowAutofit />
				</w:compat>
			</w:docPr>
      <xsl:apply-templates/>
    </w:wordDocument>
  </xsl:template>
	
  <xsl:template match="ConservationProject">
    <w:body>
			<w:p>
				<w:pPr>
					<w:pStyle w:val="Heading1" />
				</w:pPr>
				<w:r>
					<w:t>
						<xsl:value-of select="ProjectSummary/ProjectSummaryProjectName"/>
					</w:t>
				</w:r>
			</w:p>
			<w:p>
				<w:pPr>
					<w:pStyle w:val="Heading2" />
				</w:pPr>
				<w:r>
					<w:t>
						Property Information
					</w:t>
				</w:r>
			</w:p>
			<xsl:apply-templates/>
		</w:body>
  </xsl:template>

	<xsl:template match="ProjectSummary" />

	<xsl:template match="ProjectSummaryScope">
		<w:tbl>
			<w:tblPr>
				<w:tblW w:w="5000" w:type="pct"/>
				<w:tblBorders>
					<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
					<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
					<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
					<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
				</w:tblBorders>
			</w:tblPr>
			<w:tblGrid>
				<w:gridCol w:w="10296"/>
			</w:tblGrid>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Protected Area Categories
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeProtectedAreaCategoryNotes"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Legal Status
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeLegalStatus"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Legislative Status
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeLegislativeContext"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Physical Description
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopePhysicalDescription"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Biological Description
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeBiologicalDescription"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Socio-Economic Information
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeSocioEconomicInformation"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Historical Information
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeHistoricalDescription"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Cultural Description
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeCulturalDescription"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Access Information
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeAccessInformation"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Visitation Information
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeVisitationInformation"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Current Land Use
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeCurrentLandUses"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								Management Resources
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="0" w:type="auto"/>
					</w:tcPr>
					<w:p>
						<w:r>
							<w:t>
								<xsl:value-of select="ProjectSummaryScopeManagementResources"/>
							</w:t>
						</w:r>
					</w:p>
				</w:tc>
			</w:tr>
		</w:tbl>
	</xsl:template>

	<xsl:template match="ProjectResourcePool" />
	<xsl:template match="OrganizationPool" />
	<xsl:template match="ProjectSummaryLocation" />
	<xsl:template match="ProjectSummaryPlanning" />
	<xsl:template match="TncProjectData" />
	<xsl:template match="WwfProjectData" />
	<xsl:template match="WCSData" />
	<xsl:template match="RareProjectData" />
	<xsl:template match="FosProjectData" />
	<xsl:template match="BiodiversityTargetPool" />
	<xsl:template match="ConceptualModelPool" />
	<xsl:template match="ResultsChainPool" />
	<xsl:template match="DiagramFactorPool" />
	<xsl:template match="DiagramLinkPool" />
	<xsl:template match="HumanWelfareTargetPool" />
	<xsl:template match="CausePool" />
	<xsl:template match="StrategyPool" />
	<xsl:template match="ThreatReductionResultPool" />
	<xsl:template match="IntermediateResultPool" />
	<xsl:template match="GroupBoxPool" />
	<xsl:template match="TextBoxPool" />
	<xsl:template match="ScopeBoxPool" />
	<xsl:template match="KeyEcologicalAttributePool" />
	<xsl:template match="StressPool" />
	<xsl:template match="SubTargetPool" />
	<xsl:template match="GoalPool" />
	<xsl:template match="ObjectivePool" />
	<xsl:template match="IndicatorPool" />
	<xsl:template match="TaskPool" />
	<xsl:template match="ProgressReportPool" />
	<xsl:template match="ProgressPercentPool" />
	<xsl:template match="MeasurementPool" />
	<xsl:template match="AccountingCodePool" />
	<xsl:template match="FundingSourcePool" />
	<xsl:template match="BudgetCategoryOnePool" />
	<xsl:template match="BudgetCategoryTwoPool" />
	<xsl:template match="ExpenseAssignmentPool" />
	<xsl:template match="ResourceAssignmentPool" />
	<xsl:template match="ThreatRatingPool" />
	<xsl:template match="IUCNRedListSpeciesPool" />
	<xsl:template match="OtherNotableSpeciesPool" />
	<xsl:template match="AudiencePool" />
	<xsl:template match="PlanningViewConfigurationPool" />
	<xsl:template match="DeletedOrphans" />

</xsl:stylesheet>
