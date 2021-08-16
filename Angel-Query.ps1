### SalesForce Query ###
$query = "SELECT 
	Id, 
	CaseNumber,
	Priority, 
	Case_Age__c, 
    Status,
    Description,
    Preferred_Case_Language__c,
    Case_Preferred_Timezone__c,
	Tier__c,
	Entitlement_Type__c,
	Category__c, 
	Product__c, 
	Subject, 
	First_Response_Complete__c, 
	CreatedDate,
	Plan_of_Action_Status__c, 
	Case_Owner_Name__c,
    AccountId,
    IsEscalated, Escalated_Case__c,
    ClosedDate, IsClosed, isClosedText__c, 
    Account.Name,
    Account.CSM_Name__c,
    Account.CSM_Email__c,
    (SELECT CreatedDate, field, OldValue, NewValue, CreatedById 
        FROM Histories 
        WHERE CreatedDate=TODAY and field='Owner'),
    (SELECT CreatedById, body FROM Feeds),
    (SELECT MilestoneTypeId,TargetDate,TimeRemainingInDays,TimeRemainingInHrs,TimeRemainingInMins,IsViolated FROM CaseMilestones)
FROM Case 
WHERE
	RecordTypeId='012600000000nrwAAA' AND 
    ( (IsClosed=False) OR (IsClosed=True AND ClosedDate=TODAY) ) AND
	Preferred_Support_Region__c ='APAC' AND 
	Preferred_Case_Language__c != 'Japanese' AND 
    Tier__c != 'Admin'
" -replace "`n", " "
#
# (SELECT CreatedById, body FROM Feeds),
# Id IN (SELECT CaseID FROM CaseHistory ) AND Id IN (SELECT ParentId FROM CaseFeed) AND Id IN (SELECT CaseId FROM CaseMilestone)
# Account.AnnualRevenue,
# FORMAT(Account.AnnualRevenue) frmtAmnt,
# convertCurrency(Account.AnnualRevenue) cnvAmnt,
# (Status='New' or Status='Active' or Status='Re-opened') AND
# ( (IsClosed=False) OR (IsClosed=True AND ClosedDate=TODAY) ) AND
# ORDER BY Case_Owner_Name__c, Tier__c, Priority ASC, Case_Age__c DESC" -replace "`n", " "
# ORDER BY Case_Owner_Name__c, Tier__c DESC, Entitlement_Type__c DESC, Priority,Case_Age__c DESC" -replace "`n", " "
# ORDER BY Case_Owner_Name__c, Tier__c DESC, Account.CSM_Name__c DESC, Priority,Case_Age__c DESC" -replace "`n", " "
# ORDER BY Case_Owner_Name__c, Priority,Case_Age__c DESC" -replace "`n", " "

$query = "SELECT 
	Id, 
	CaseNumber,
	Priority, 
	Case_Age__c, 
    Status,
    Description,
    Preferred_Case_Language__c,
    Case_Preferred_Timezone__c,
	Tier__c,
	Entitlement_Type__c,
	Category__c, 
	Product__c, 
	Subject, 
	First_Response_Complete__c, 
	CreatedDate,
	Plan_of_Action_Status__c, 
	Case_Owner_Name__c,
    AccountId,
    IsEscalated, Escalated_Case__c,
    ClosedDate, IsClosed, isClosedText__c, 
    Account.Name,
    Account.CSM_Name__c,
    Account.CSM_Email__c,
    (SELECT CreatedDate, field, OldValue, NewValue, CreatedById 
        FROM Histories 
    WHERE CreatedDate=TODAY and field='Owner'),
    (SELECT CreatedById, body FROM Feeds)
FROM Case 
WHERE
	RecordTypeId='012600000000nrwAAA' AND 
    ( (IsClosed=False) OR (IsClosed=True AND ClosedDate=TODAY) ) AND
	Preferred_Support_Region__c ='APAC' AND 
	Preferred_Case_Language__c != 'Japanese' AND 
    Tier__c != 'Admin'
" -replace "`n", " "


# Removed Protip...
$query = "SELECT 
	Id, 
	CaseNumber,
	Priority, 
	Case_Age__c, 
    Status,
    Description,
    Preferred_Case_Language__c,
    Case_Preferred_Timezone__c,
	Tier__c,
	Entitlement_Type__c,
	Category__c, 
	Product__c, 
	Subject, 
	First_Response_Complete__c, 
	CreatedDate,
	Plan_of_Action_Status__c, 
	Case_Owner_Name__c,
    AccountId,
    IsEscalated, Escalated_Case__c,
    ClosedDate, IsClosed, isClosedText__c, 
    Account.Name,
    Account.CSM_Name__c,
    Account.CSM_Email__c,
    (SELECT CreatedDate, field, OldValue, NewValue, CreatedById 
        FROM Histories 
    WHERE CreatedDate=TODAY and field='Owner')
FROM Case 
WHERE
	RecordTypeId='012600000000nrwAAA' AND 
    ( (IsClosed=False) OR (IsClosed=True AND ClosedDate=TODAY) ) AND
	Preferred_Support_Region__c ='APAC' AND 
	Preferred_Case_Language__c != 'Japanese' AND 
    Tier__c != 'Admin'
" -replace "`n", " "

