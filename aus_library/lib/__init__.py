import openstudio
import pandas as pd
from pathlib import Path
from inspect import getsourcefile
from os.path import abspath
import json



def sheet_to_json(sheet_name):
    path_to_file = Path(abspath(getsourcefile(lambda:0))).parent.joinpath("resources/resources.xlsx").resolve()
    sheet = pd.read_excel(path_to_file, sheet_name=sheet_name)
    jason = json.loads( sheet.to_json(orient='records') )
    return jason

def read_excel():

    sheets = ["people", "schedules", "materials","lights", "equipment", "infiltration", "outdoor_air", "space_types"]

    return {k: sheet_to_json(k) for k in sheets}


def create_complex_schedule(model, options = {}):
    defaults = {
        'name': None,
        'default_day': ['always_on', [24.0, 1.0]]
    }
    # merge user inputs with defaults
    options = {**defaults, **options}
    # ScheduleRuleset
    sch_ruleset = openstudio.model.ScheduleRuleset(model)
    if options['name']:
        sch_ruleset.setName(options['name'])
    # Winter Design Day
    if options['winter_design_day'] is not None:
        winter_dsn_day = openstudio.model.ScheduleDay(model)
        sch_ruleset.setWinterDesignDaySchedule(winter_dsn_day)
        winter_dsn_day = sch_ruleset.winterDesignDaySchedule()
        winter_dsn_day.setName(f"{sch_ruleset.name()} Winter Design Day")
        for data_pair in options['winter_design_day']:
            hour = int(data_pair[0])
            min = int((data_pair[0] - hour) * 60)
            winter_dsn_day.addValue(openstudio.Time(0, hour, min, 0), data_pair[1])
    # Summer Design Day
    if options['summer_design_day'] is not None:
        summer_dsn_day = openstudio.Model.ScheduleDay(model)
        sch_ruleset.setSummerDesignDaySchedule(summer_dsn_day)
        summer_dsn_day = sch_ruleset.summerDesignDaySchedule()
        summer_dsn_day.setName(f"{sch_ruleset.name()} Summer Design Day")
        for data_pair in options['summer_design_day']:
            hour = int(data_pair[0])
            min = int((data_pair[0] - hour) * 60)
            summer_dsn_day.addValue(openstudio.Time(0, hour, min, 0), data_pair[1])
    # Default Day
    default_day = sch_ruleset.defaultDaySchedule()
    default_day.setName(f"{sch_ruleset.name()} {options['default_day'][0]}")
    default_data_array = options['default_day']
    del default_data_array[0]
    for data_pair in default_data_array:
        hour = int(data_pair[0])
        min = int((data_pair[0] - hour) * 60)
        default_day.addValue(openstudio.Time(0, hour, min, 0), data_pair[1])
    # Rules
    if options['rules'] is not None:
        for data_array in options['rules']:
            rule = openstudio.Model.ScheduleRule(sch_ruleset)
            rule.setName(f"{sch_ruleset.name()} {data_array[0]} Rule")
            date_range = data_array[1].split('-')
            start_date = date_range[0].split('/')
            end_date = date_range[1].split('/')
            rule.setStartDate(model.getYearDescription().makeDate(int(start_date[0]), int(start_date[1])))
            rule.setEndDate(model.getYearDescription().makeDate(int(end_date[0]), int(end_date[1])))
            days = data_array[2].split('/')
            if 'Sun' in days:
                rule.setApplySunday(True)
            if 'Mon' in days:
                rule.setApplyMonday(True)
            if 'Tue' in days:
                rule.setApplyTuesday(True)
            if 'Wed' in days:
                rule.setApplyWednesday(True)
            if 'Thu' in days:
                rule.setApplyThursday(True)
            if 'Fri' in days:
                rule.setApplyFriday(True)
            if 'Sat' in days:
                rule.setApplySaturday(True)
            day_schedule = rule.daySchedule()
            day_schedule.setName(f"{sch_ruleset.name()} {data_array[0]}")
            del data_array[0:3]
            for data_pair in data_array:
                hour = int(data_pair[0])
                min = int((data_pair[0] - hour) * 60)
                day_schedule.addValue(openstudio.Time(0, hour, min, 0), data_pair[1])
    return sch_ruleset

def make_schedule_sets():
    return "lights"

def create_people_load(osm, peoples:dict):
    # make a people def
    
    def _make(data:dict):
        people_def = openstudio.model.PeopleDefinition(osm)
        people_def.setSpaceFloorAreaperPerson( data.get("area per person") )
        people_def.setName( data.get("description") )
        #setSpaceFloorAreaperPerson
        #setNumberofPeople
        #setFractionRadiant
        #setSensibleHeatFraction

        # make a people load
        people_load = openstudio.model.People(people_def)
        #set work efficiency schedule name
        #set clothing insulation schedule name 
        #set air velocity schedule name

    list(map( _make, peoples))

def create_lights_load(osm, lights:dict):
    
    def _make(data:dict):
        lights_def = openstudio.model.LightsDefinition(osm)
        lights_def.setWattsperSpaceFloorArea( data.get("Adjusted IPD") )
        lights_load = openstudio.model.Lights(lights_def)

    list(map( _make, lights))


def create_electric_equipment_load(osm, equipments:dict):

    def _make(data:dict):
        equip_def = openstudio.model.EquipmentDefinition(osm)
        equip_def.setWattsperSpaceFloorArea( data.get("W/m2") )
        equip_load = openstudio.model.Lights(equip_def)

    list(map( _make, equipments))

def create_infiltration_objects(osm, infiltrations:dict):

    def _make(data:dict):
        infiltration = openstudio.model.SpaceInfiltrationDesignFlowRate(osm)
        infiltration.setDesignFlowRateCalculationMethod("AirChanges/Hour")
        infiltration.setAirchangesPerHour( infiltration.get("hvac_off") )

    list(map( _make, infiltrations))


def create_outdoor_air_objects(osm, outdoor_airs:dict):
    def _make(data:dict):
        ventilation = openstudio.model.DesignSpecificationOutdoorAir(osm)
        ventilation.setOutdoorAirFlowperPerson( outdoor_airs.get("L/s/person") )

    list(map( _make, outdoor_airs))

def create_space_types():
    pass

