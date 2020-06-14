# Rebuild-SfBUser.ps1
Enables users for SfB, sets their lineURI and assigns (grants) policies

## BASED ON

File Name: EnableLyncusers.ps1<br>
Version: 0.4<br>
Last Update: 17-May-2014<br>
Author: Paul Bloem, http://ucsorted.com

## Overview

This script will run through a csv input file and generate the required commands to enable users for SfB & assign (grant) policies.

It's great when you're fault-finding and need to blow away and rebuild Lab users. 

The csv file column headers are used as variables in the script, so they need to be unchanged.

[FUTURE] The script can be set to choose any of the usual identifying column headers as the "-Identity" of the user being created:
- DN
- UPN
- SAM

### Format

The csv file will need the following column headers as a minimum:-
- SamAccountName
- SIPAddress
- RegistrarPool

If any policy is to be left as default then simply edit out the policy in the script. The provided labels provide for easy recognition within the script.

Make sure that the data within each row matches already configured, valid policies. (If you're exporting, disabling and then rebuilding them, this shouldn't be an issue for you).

### Example file structure

```powerShell
SamAccountName,SIPAddress,LineUri,LocationPolicy,DialPlanPolicy,VoicePolicy,ConferencingPolicy,ExternalPolicy,ClientPolicy,MobilityPolicy,RegistrarPool
PanP,peter.pan@SfBsorted.co.nz,+6499702700,Auckland,Auckland_Dial_Plan,NZ_International,Audio_Only,Full External_Access,Photo_Control,SfBpool01.SfBsorted.co.nz
```

### Create from existing

```powerShell
get-csuser -filter {RegistrarPool -ne $null} | select SamAccountName,SipAddress,RegistrarPool,DialPlan,LineURI,EnterpriseVoiceEnabled,*policy | export-csv -NoTypeInformation -path <FILENAME.CSV>
```
	
&nbsp;<br>
\- Greig.
