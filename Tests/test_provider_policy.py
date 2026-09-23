from datetime import date, time
from rehab_core.provider_policy import *
D=date(2026,9,24); T1030=time(10,30); T1045=time(10,45)
def b(pid='Θ1'):
    x=ProviderPolicyBook(); x.profiles[pid]=ProviderProfile(pid,pid); return x

def test_default_therapist_capacity_is_six(): assert b().effective_state('Θ1',D).max_slots==6
def test_student_capacity_is_five():
    x=ProviderPolicyBook(); x.profiles['Φ1']=ProviderProfile('Φ1','Φ1',ProviderType.STUDENT); assert x.effective_state('Φ1',D).max_slots==5
def test_temporary_max_override_wins():
    x=b(); x.overrides.append(ProviderOverride('Θ1',OverrideKind.MAX_LOAD,D,D,max_load=3)); assert x.effective_state('Θ1',D).max_slots==3
def test_exclude_from_replacements():
    x=b(); x.overrides.append(ProviderOverride('Θ1',OverrideKind.EXCLUDE_REPLACEMENTS,D,D)); assert x.can_take_replacement(provider_id='Θ1',on_date=D,requested_time=T1045,occupied_times=[]).reason=='excluded_from_replacements'
def test_blocked_time_is_hard_stop():
    x=b(); x.overrides.append(ProviderOverride('Θ1',OverrideKind.BLOCK_TIME,D,D,time_slot=T1045)); assert x.can_take_replacement(provider_id='Θ1',on_date=D,requested_time=T1045,occupied_times=[]).reason=='time_blocked'
def test_manager_zero_base_load():
    x=b(); x.leadership.append(LeadershipAssignment('Θ1',LeadershipRole.MANAGER,D)); st=x.effective_state('Θ1',D); assert st.max_slots==0 and not st.replacement_eligible
def test_manager_only_explicit_replacement_time():
    x=b(); x.leadership.append(LeadershipAssignment('Θ1',LeadershipRole.MANAGER,D)); x.overrides.append(ProviderOverride('Θ1',OverrideKind.OPEN_REPLACEMENT_TIME,D,D,time_slot=T1045)); assert not x.can_take_replacement(provider_id='Θ1',on_date=D,requested_time=T1030,occupied_times=[]).allowed; assert x.can_take_replacement(provider_id='Θ1',on_date=D,requested_time=T1045,occupied_times=[]).allowed
def test_acting_manager_takes_precedence():
    x=b('ΠΡ'); x.profiles['ΑΝ']=ProviderProfile('ΑΝ','ΑΝ'); x.leadership=[LeadershipAssignment('ΠΡ',LeadershipRole.MANAGER,date(2026,1,1)),LeadershipAssignment('ΑΝ',LeadershipRole.ACTING_MANAGER,D,date(2026,9,26))]; assert x.active_leader(D).provider_id=='ΑΝ'
