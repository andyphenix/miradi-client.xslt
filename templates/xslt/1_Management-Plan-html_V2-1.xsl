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
	xmlns:w="http://schemas.microsoft.com/office/word/2010/wordml" 
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:v="urn:schemas-microsoft-com:vml"
	xmlns:html="http://www.w3.org/TR/REC-html40"
	xmlns:w10="urn:schemas-microsoft-com:office:word" 
	xmlns:sl="http://schemas.microsoft.com/schemaLibrary/2003/core" 
	xmlns:aml="http://schemas.microsoft.com/aml/2001/core" 
	xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint" 
	xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882">
	
	 


	
	<!-- Provide reference to local directory for images-->
	<!-- Images must be stored in directory located in specific path named below -->
	<!-- For Windows, use file:///C:/BHA/images/ -->
	<!-- For Mac, use /BHA/images/ -->
	<!-- CURRENTLY SHOWING WINDOWS PATH CHANGE IF USING MAC-->
	<xsl:variable name="imgPath">file:///C:/BHA/images/</xsl:variable>
	
	

	<xsl:template match="ConservationProject">
		<html>
			<head>
				<style>
					body { font-family: Arial; font-size: 10pt; margin: 0;}
					td { vertical-align: top;}
					td b { font-size: 14pt; color: #4F81BD; }
					td.header {vertical-align: bottom;} 
					.break { page-break-before: always; }
					.conservationtargetpicture { width: 5cm; }
					.coverpagedate { font-size: 14pt; text-align:center; }
					.coverpagepicture { text-align: left;}
					.coverpagetitle { font-size: 24pt; font-weight: bold; }
					.instruction { font-style: italic; }
					.narrow {font-family: Arial Narrow}
					.subheading { color: #4F81BD; font-weight: bold; }
					.toc {font-size: 13pt; font-weight: bold; color: #4F81BD; }
					.vision { font-style: italic; }
					.targetImg {padding: 2px;}
					.health-vg { background-color: #008000; }
					.health-g { background-color: #00ff00; }
					.health-f {background-color: #ffff00;}
					.health-p {background-color: #ff0000;}
					img.rc-img { vertical-align:middle; height: 5in;}
			
				</style>
			</head>
			<body>
				<xsl:apply-templates select="ProjectSummary" mode="titlepage" />
				<xsl:apply-templates select="ProjectSummary" mode="introduction"/>
				<xsl:apply-templates select="ProjectSummaryScope" mode="bioregionalcontext" />
				<xsl:apply-templates select="ProjectSummaryLocation" />
				<xsl:apply-templates select="ProjectSummaryScope" mode="legalstatus" />
				<xsl:apply-templates select="ProjectSummaryScope" mode="history" />
				<xsl:apply-templates select="ProjectSummaryScope" mode="culture" />
				<xsl:apply-templates select="ProjectSummaryScope" mode="vision" />
				<xsl:apply-templates select="ConceptualModelPool" />
				<xsl:apply-templates select="BiodiversityTargetPool" mode="targets" />
				<xsl:apply-templates select="BiodiversityTargetPool" mode="viability" />
				<xsl:apply-templates select="BiodiversityTargetPool" mode="threats" />
				<xsl:apply-templates select="ResultsChainPool" />
				<xsl:apply-templates select="BiodiversityTargetPool" mode="progress" />
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
		<p class="coverpagetitle"><xsl:value-of select="ProjectSummaryProjectName"/> Property Plan</p>
		<br />
		<br />
		<!--<br />
		<br />
		<br />
		<br /> -->
		<div class="coverpagepicture">
			<p class="instruction"><img src="{$imgPath}CoverPageImage.png" width="700" height="525" /></p>
		</div>
		<br />
		<!--<br />
		<br />
		<br />
		<br />
		<br />-->
		<div class="coverpagedate">
			<xsl:value-of select="ProjectSummaryDataEffectiveDate"/>
		</div>
		<div class="break">
			<span class="toc">Table of Contents</span>
		</div>
		<p class="instruction">
			Insert table of contents (in Word: References, Table of Contents)
		</p>
		<h3>Acknowledgements and Status</h3>
		<p class="narrow">
			<!--<xsl:value-of select="ProjectSummaryProjectStatus"/> -->
			
			<xsl:call-template name="PreserveLineBreaks">
				<xsl:with-param name="text" select="ProjectSummaryProjectStatus"/>
			</xsl:call-template>
		</p>
		
		<p class="narrow">
			<xsl:value-of select="ProjectSummaryNextSteps"/>
		</p>
	</xsl:template>

	<xsl:template match="ProjectSummary" mode="introduction">
		<h2 class="break">
			Introduction
		</h2>
		<p>
			
			<xsl:call-template name="PreserveLineBreaks">
				<xsl:with-param name="text" select="ProjectSummaryProjectDescription"/>
			</xsl:call-template>
			
			<!-- <xsl:value-of select="ProjectSummaryProjectDescription" /> -->
		</p>
	</xsl:template>
	
	<xsl:template match="ProjectSummaryScope" mode="bioregionalcontext">
		<h2>
			Bioregional and Landscape Context
		</h2>
		<p>
			
			<xsl:call-template name="PreserveLineBreaks">
				<xsl:with-param name="text" select="ProjectSummaryScopeProjectAreaNote"/>
			</xsl:call-template>
			
			<!--<xsl:value-of select="ProjectSummaryScopeProjectAreaNote"/>-->
		</p>
	</xsl:template>

	<xsl:template match="ProjectSummaryLocation">
		<h2>
			Location Details
		</h2>
		<p>
			<xsl:call-template name="PreserveLineBreaks">
				<xsl:with-param name="text" select="ProjectSummaryLocationLocationDetail"/>
			</xsl:call-template>
			
			<!--<xsl:value-of select="ProjectSummaryLocationLocationDetail"/>-->
		</p>
		<p class="instruction">
			<img src="{$imgPath}ProjectMap.png" width="700" height="525"/>
		</p>
	</xsl:template>

	<xsl:template match="ProjectSummaryScope" mode="legalstatus">
		<h2>
			Legal Status
		</h2>
		<p>
			
			<xsl:call-template name="PreserveLineBreaks">
				<xsl:with-param name="text" select="ProjectSummaryScopeLegalStatus"/>
			</xsl:call-template>
			
			<!--<xsl:value-of select="ProjectSummaryScopeLegalStatus"/>-->
		</p>
		<p>
			<span class="subheading" xml:space="preserve">Protected Area Category: </span>
			
			<xsl:call-template name="PreserveLineBreaks">
				<xsl:with-param name="text" select="ProjectSummaryScopeProtectedAreaCategoryNotes"/>
			</xsl:call-template>
			
			<!--<xsl:value-of select="ProjectSummaryScopeProtectedAreaCategoryNotes" />-->
		</p>
		
	</xsl:template>

	<xsl:template match="ProjectSummaryScope" mode="history">
		<h2>
			History
		</h2>
		<p>
			
			<xsl:call-template name="PreserveLineBreaks">
				<xsl:with-param name="text" select="ProjectSummaryScopeHistoricalDescription"/>
			</xsl:call-template>
			
			<!--<xsl:value-of select="ProjectSummaryScopeHistoricalDescription"/>-->
		</p>
	</xsl:template>

	<xsl:template match="ProjectSummaryScope" mode="culture">
		<h2>
			Cultural Description
		</h2>
		<p>
			<xsl:call-template name="PreserveLineBreaks">
				<xsl:with-param name="text" select="ProjectSummaryScopeCulturalDescription"/>
			</xsl:call-template>
			
			<!--<xsl:value-of select="ProjectSummaryScopeCulturalDescription"/>-->
		</p>
	</xsl:template>

	<xsl:template match="ProjectSummaryScope" mode="vision">
		<h2>
			Scope
		</h2>
		<p>
			<xsl:call-template name="PreserveLineBreaks">
				<xsl:with-param name="text" select="ProjectSummaryScopeProjectScope"/>
			</xsl:call-template>
			
			<!--<xsl:value-of select="ProjectSummaryScopeProjectScope"/>-->
		</p>
		<h2>
			Our Conservation Vision
		</h2>
		<span class="vision">
			
			<xsl:call-template name="PreserveLineBreaks">
				<xsl:with-param name="text" select="ProjectSummaryScopeProjectVision"/>
			</xsl:call-template>
			
			<!--<xsl:value-of select="ProjectSummaryScopeProjectVision"/>-->
		</span>
	</xsl:template>

	<xsl:template match="ConceptualModelPool">
		<xsl:for-each select="ConceptualModel">
		<h2 class="break">
			<xsl:value-of select="ConceptualModelName"/>
		</h2>
		<p>
			
			<xsl:if test="ConceptualModelDetails">
				<xsl:value-of select="ConceptualModelDetails"/>
			</xsl:if>
		</p>
		
			<xsl:variable name="cm-id" select="@Id"/>
			<p><img src="{$imgPath}CM-{$cm-id}.png" width="700" height="525" alt="{$cm-id}"/>
			</p>
		</xsl:for-each>
		
		<img src="{$imgPath}Miradi_Legend.png" width="700" height="91"/>
	</xsl:template>

	<xsl:template match="BiodiversityTargetPool" mode="targets">
		<h2 class="break">
			Our Conservation Targets
		</h2>
		<table cellpadding="3" cellspacing="0">
			<xsl:for-each select="BiodiversityTarget">
				<xsl:sort select="BiodiversityTargetId" />
				<xsl:variable name="bt-id" select="@Id"/>
				<tr>
					<!--<td align="left">
						<span class="instruction">
							<img src="{$imgPath}Target{$bt-id}.png" align="left"/>
						</span>
					</td> -->
					<td>
						
						<p>
							<b>
								<xsl:call-template name="PreserveLineBreaks">
									<xsl:with-param name="text" select="BiodiversityTargetName"/>
								</xsl:call-template>
								
								<!--<xsl:value-of select="BiodiversityTargetName"/>-->
							</b>
						</p>
						<br/>
						<div>
							<img src="{$imgPath}Target{$bt-id}.png" align="left" hspace="10" vspace="10" width="400"/>
						</div>
						
						
							<xsl:call-template name="PreserveLineBreaks">
								<xsl:with-param name="text" select="BiodiversityTargetDetails"/>
							</xsl:call-template>
							
							<!--<xsl:value-of select="BiodiversityTargetDetails"/>-->
						
						<xsl:call-template name="Goals" />
						<br />
					</td>
				</tr>
			</xsl:for-each>
		</table>
	</xsl:template>

	<xsl:template match="BiodiversityTargetPool" mode="viability">
		<h2 class="break">
			Viability Assessment
		</h2>
		<p>
			Viability analysis looks at each of the conservation targets to determine how to measure its "health" over time, and then to identify how the target is doing today and what a "healthy state" might look like. This helps to identify which of our targets are most in need of immediate attention, and for measuring success over time. The Status ratings are derived by rolling up measurements for underlying indicators in our Monitoring Plan (described later).
		</p>
		<p>
			<xsl:call-template name="ViabilityAssessment" />
		</p>
		<p class="narrow">
			<b>Legend: </b>
			<b>Very Good - </b>Ecologically desirable status; requires little intervention for maintenance.;
			<b>Good - </b>Indicator within acceptable range of variation; some intervention required for maintenance;
			<b>Fair - </b>Outside acceptable range of variation; requires human intervention;
			<b>Poor - </b>Restoration increasingly difficult; may result in extirpation of target.
		</p>
	</xsl:template>

	<xsl:template match="BiodiversityTargetPool" mode="threats">
		<h2 class="break">
			The key Threats to our Targets
		</h2>
		<p>
			The key threats to each of our targets have been assessed to help identify what action is required. The Threat ratings are based on a combination of scope (the area affected by the threat), severity (the degree of damage) and irreversibility (how permanent the damage is).  These ratings show our estimates of the Threats at the last major review of the plan; implementation of our strategies is aimed at reducing the most critical threats.
		</p>
		<p class="instruction">
			<a href="{$imgPath}ThreatRatingTable.rtf" target="_blank"><img src="{$imgPath}ThreatRatingTable.png" width="700" height="525"/></a>
		</p>
		<xsl:call-template name="VeryHighThreatsList" />
	</xsl:template>

	<xsl:template name="Goals">
		<xsl:for-each select="BiodiversityTargetGoalIds/GoalId">
			<xsl:call-template name="Goal">
				<xsl:with-param name="GoalParam">
					<xsl:value-of select="text()"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="Goal">
		<xsl:param name="GoalParam" />
		<p>
			<strong>Goal: </strong>
			<xsl:value-of select="/ConservationProject/GoalPool/Goal[@Id=$GoalParam]/GoalName" />
			<xsl:value-of select="/ConservationProject/GoalPool/Goal[@Id=$GoalParam]/GoalDetails" />
		</p>
	</xsl:template>

	<xsl:template name="ViabilityAssessment">
		<table border="1" cellpadding="5" cellspacing="0">
			<tr>
				<td class="header">
					<h4>
						Target
					</h4>
				</td>
				<td class="header">
					<h4>
						Current Status
					</h4>
				</td>
				<td class="header">
					<h4>
						Desired Status
					</h4>
				</td>
				<td class="header">
					<h4>
						Details
					</h4>
				</td>
			</tr>
			<xsl:for-each select="BiodiversityTarget">
				<xsl:sort select="BiodiversityTargetId" />
				<tr>
					<td>
						<strong>
							<xsl:value-of select="BiodiversityTargetName"/>
						</strong>
					</td>
					<td>
						<xsl:attribute name="class">
							<xsl:choose>
								<xsl:when test="BiodiversityTargetViabilityStatus = 1">health-p</xsl:when>
								<xsl:when test="BiodiversityTargetViabilityStatus = 2">health-f</xsl:when>
								<xsl:when test="BiodiversityTargetViabilityStatus = 3">health-g</xsl:when>
								<xsl:when test="BiodiversityTargetViabilityStatus = 4">health-vg</xsl:when>
							</xsl:choose>
						</xsl:attribute>
						<xsl:if test="BiodiversityTargetViabilityStatus = 1">
							Poor
						</xsl:if>
						<xsl:if test="BiodiversityTargetViabilityStatus = 2">
							Fair
						</xsl:if>
						<xsl:if test="BiodiversityTargetViabilityStatus = 3">
							Good
						</xsl:if>
						<xsl:if test="BiodiversityTargetViabilityStatus = 4">
							Very Good
						</xsl:if>
					</td>
					<td>
						<xsl:attribute name="class">
							<xsl:choose>
								<xsl:when test="BiodiversityTargetViabilityDesired = 1">health-p</xsl:when>
								<xsl:when test="BiodiversityTargetViabilityDesired = 2">health-f</xsl:when>
								<xsl:when test="BiodiversityTargetViabilityDesired = 3">health-g</xsl:when>
								<xsl:when test="BiodiversityTargetViabilityDesired = 4">health-vg</xsl:when>
							</xsl:choose>
						</xsl:attribute>
						<xsl:if test="BiodiversityTargetViabilityDesired = 1">
							Poor
						</xsl:if>
						<xsl:if test="BiodiversityTargetViabilityDesired = 2">
							Fair
						</xsl:if>
						<xsl:if test="BiodiversityTargetViabilityDesired = 3">
							Good
						</xsl:if>
						<xsl:if test="BiodiversityTargetViabilityDesired = 4">
							Very Good
						</xsl:if>
					</td>
					<td>
						<xsl:call-template name="PreserveLineBreaks">
							<xsl:with-param name="text" select="BiodiversityTargetCurrentStatusJustification"/>
						</xsl:call-template>
						
						<!--<xsl:value-of select="BiodiversityTargetCurrentStatusJustification"/>-->
					</td>
				</tr>
			</xsl:for-each>
		</table>
	</xsl:template>

	<xsl:template name="VeryHighThreatsList">
		<xsl:for-each select="BiodiversityTarget">
			<xsl:sort select="BiodiversityTargetId" />
			<xsl:call-template name="VeryHighThreatTarget">
				<xsl:with-param name="TargetParam">
					<xsl:value-of select="@Id"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="VeryHighThreatTarget">
		<xsl:param name="TargetParam" />
		<xsl:for-each select="/ConservationProject/ThreatRatingPool/ThreatRating">
			<xsl:if test="ThreatRatingTargetId/BiodiversityTargetId = $TargetParam">
				<xsl:if test="ThreatRatingRatings/SimpleThreatRating/SimpleThreatRatingSeverity = 4 and ThreatRatingRatings/SimpleThreatRating/SimpleThreatRatingScope = 4">
					<xsl:if test="ThreatRatingRatings/SimpleThreatRating/SimpleThreatRatingIrreversibility = 3 or ThreatRatingRatings/SimpleThreatRating/SimpleThreatRatingIrreversibility = 4">
						<p>
						<b>
							<xsl:value-of select="/ConservationProject/BiodiversityTargetPool/BiodiversityTarget[@Id=$TargetParam]/BiodiversityTargetName"/>:
							<xsl:variable name="ThreatId">
								<xsl:value-of select="ThreatRatingThreatId/ThreatId"/>
							</xsl:variable>
							<!--<w:t xml:space="preserve"><xsl:value-of select="/ConservationProject/CausePool/Cause[@Id=$ThreatId]/CauseName"/>: </w:t> -->
							
							<xsl:value-of select="/ConservationProject/CausePool/Cause[@Id=$ThreatId]/CauseName"/>:
							
						</b>
							<xsl:value-of select="ThreatRatingComments"/>
					</p>
					</xsl:if>
				</xsl:if>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>

	<xsl:template match="ResultsChainPool">
		<h2 class="break">
			Our Strategies
		</h2>
		<p>
			The following pages describe the Strategies we will pursue to reduce Threats and protect the Targets. Each strategy is described in text, and then depicted in a Results Chain diagram which shows our assumptions and the logic that we expect to play out over time (refer Legend below). The diagrams show the Strategy and the associated Activities that we will perform. Execution of these Activities should lead to one of more Intermediate Results, which should ultimately result in reduction of the Threat , which in turn will improve the condition of our Targets . Objectives have been set for each Strategy and Indicators have been defined to help monitor our progress. We will continue to monitor these Indicators, and adjust our Strategy if we are not achieving the expected outcomes.
		</p>
		 
		 	<xsl:for-each select="ResultsChain">
		 		<xsl:sort select="ResultsChainId" />
		 		<xsl:variable name="rc-id" select="@Id"/>
		 		<table cellspacing="0" cellpadding="3">
		 		<tr>
		 			<td>
		 				<h3>
		 					<xsl:attribute name="class">
		 						<xsl:if test="position()!= 1">break</xsl:if>
		 					</xsl:attribute>
		 					<xsl:value-of select="ResultsChainName"/>
		 				</h3>
		 				<p>
		 					<xsl:call-template name="PreserveLineBreaks">
		 						<xsl:with-param name="text" select="ResultsChainDetails"/>
		 					</xsl:call-template>
		 				</p>
		 				
		 				
		 				
		 				
		 			</td>
		 		</tr>
		 		</table>
		 		
		 		<p>
		 			<div id="block">
		 				<img src="{$imgPath}RC-{$rc-id}.png" />
		 			</div>
		 		</p>
		 		
		 	</xsl:for-each>
		 
		
	</xsl:template>
	
	<xsl:template match="BiodiversityTargetPool" mode="progress">
		<h2 class="break">
			Measuring Our Progress
		</h2>
		<p>
			The fundamental question facing any team is: "Are our strategies working?" To answer this question, we will periodically collect data on a number of indicators that gauge how well our strategies are keeping the critical threats in check and, in turn, whether the viability of our conservation targets is improving. A monitoring framework for the viability of our targets is outlined below.
		</p>
		<table>
			<xsl:for-each select="BiodiversityTarget">
				<xsl:sort select="BiodiversityTargetId" />
				<tr>
					<td>
						<h4>
							<xsl:value-of select="BiodiversityTargetName"/>
						</h4>
					</td>
				</tr>
				<xsl:call-template name="KeyAttributes" />
			</xsl:for-each>
		</table>
	</xsl:template>

	<xsl:template name="KeyAttributes">
		<tr>
			<td>
				<strong>
					Key attribute
				</strong>
			</td>
			<td>
				<strong>
					Indicator
				</strong>
			</td>
		</tr>
		<xsl:for-each select="BiodiversityTargetKeyEcologicalAttributeIds/KeyEcologicalAttributeId">
			<tr>
				<td>
					<xsl:call-template name="KeyEcologicalAttributeName">
						<xsl:with-param name="KeyEcologicalAttributeId">
							<xsl:value-of select="text()"/>
						</xsl:with-param>
					</xsl:call-template>
				</td>
				<td>
					<xsl:call-template name="KeyEcologicalAttributeIndicators">
						<xsl:with-param name="KeyEcologicalAttributeId">
							<xsl:value-of select="text()"/>
						</xsl:with-param>
					</xsl:call-template>
				</td>
			</tr>
			<tr>
				<td>
					
				</td>
			</tr>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="KeyEcologicalAttributeName">
		<xsl:param name="KeyEcologicalAttributeId" />
		<xsl:value-of select="/ConservationProject/KeyEcologicalAttributePool/KeyEcologicalAttribute[@Id=$KeyEcologicalAttributeId]/KeyEcologicalAttributeName"/>
	</xsl:template>

	<xsl:template name="KeyEcologicalAttributeIndicators">
		<xsl:param name="KeyEcologicalAttributeId" />
		<xsl:if test="count(/ConservationProject/KeyEcologicalAttributePool/KeyEcologicalAttribute[@Id=$KeyEcologicalAttributeId]/KeyEcologicalAttributeIndicatorIds/IndicatorId) = 0">
			<br />
		</xsl:if>
		<xsl:for-each select="/ConservationProject/KeyEcologicalAttributePool/KeyEcologicalAttribute[@Id=$KeyEcologicalAttributeId]/KeyEcologicalAttributeIndicatorIds">
			<xsl:call-template name="KeyEcologicalAttributeIndicator">
				<xsl:with-param name="IndicatorId">
					<xsl:value-of select="IndicatorId"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="KeyEcologicalAttributeIndicator">
		<xsl:param name="IndicatorId" />
		<xsl:value-of select="/ConservationProject/IndicatorPool/Indicator[@Id=$IndicatorId]/IndicatorName"/>
	</xsl:template>
	
	<!-- preserving line breaks added 2011-07-20 -->
	<xsl:template name="PreserveLineBreaks">
		<xsl:param name="text"/>
		<xsl:choose>
			<xsl:when test="contains($text,'&#xA;')">
				<xsl:value-of select="substring-before($text,'&#xA;')"/>
				<br/>
				<xsl:call-template name="PreserveLineBreaks">
					<xsl:with-param name="text">
						<xsl:value-of select="substring-after($text,'&#xA;')"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$text"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

</xsl:stylesheet>
