import typing
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
    sheets = ["people", "schedules","schedule_sets", "materials","lights", "equipment", "infiltration", "outdoor_air", "space_types"]
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
        lights_def.setName( data.get("description") )
        

    list(map( _make, lights))

def create_electric_equipment_load(osm, equipments:dict):
    def _make(data:dict):
        equip_def = openstudio.model.ElectricEquipmentDefinition(osm)
        equip_def.setWattsperSpaceFloorArea( data.get("W/m2") )
        equip_def.setName( data.get("description") )
        equip_load = openstudio.model.ElectricEquipment(equip_def)

    list(map( _make, equipments))

def create_infiltration_objects(osm, infiltrations:dict):

    def _make(data:dict):
        infiltration = openstudio.model.SpaceInfiltrationDesignFlowRate(osm)
        infiltration.setName(data.get("description"))
        infiltration.setAirChangesperHour( data.get("hvac_off") )

    list(map( _make, infiltrations))

def create_outdoor_air_objects(osm, outdoor_airs:dict):
    def _make(data:dict):
        ventilation = openstudio.model.DesignSpecificationOutdoorAir(osm)
        ventilation.setOutdoorAirFlowperPerson( data.get("L/s/person") )
        ventilation.setName(data.get("description"))

    list(map( _make, outdoor_airs))

def create_space_types(osm, space_types_dicts:list):
    
    def _make(data:dict):
        space_type = openstudio.model.SpaceType(osm)
        space_type.setName( data.get("name"))

        lights_def = openstudio.model.getLightsDefinitionByName(osm, data.get("lights") ).get()
        lights = openstudio.model.Lights(lights_def)
        space_load = openstudio.model.SpaceLoad(lights)
        space_load.setSpaceType( space_type)

        equip_def = openstudio.model.getElectricEquipmentDefinitionByName(osm, data.get("equipment") ).get()
        equip = openstudio.model.ElectricEquipment(equip_def)
        space_load = openstudio.model.SpaceLoad(equip)
        space_load.setSpaceType( space_type)

        people_def = openstudio.model.getPeopleDefinitionByName(osm, data.get("people") ).get()
        people = openstudio.model.People(people_def)
        space_load = openstudio.model.SpaceLoad(people)
        space_load.setSpaceType( space_type)

        inf = openstudio.model.getSpaceInfiltrationDesignFlowRateByName(osm, data.get("infiltration") ).get()
        space_load = openstudio.model.SpaceLoad(inf)
        space_load.setSpaceType( space_type)

        s = osm.getDefaultScheduleSetByName("Office").get()
        space_type.setDefaultScheduleSet(s)
        oa = osm.getDesignSpecificationOutdoorAirByName("office 7.5 L/s/person").get()
        space_type.setDesignSpecificationOutdoorAir( oa)


    list(map( _make, space_types_dicts))

    pass

def set_in(d, path, value):
    parts = path.split('.') if isinstance(path, str) else path
    if len(parts) <= 1:
        d[parts[0]] = value
        return d
    if not parts[0] in d:
        d[parts[0]] = {}
    set_in(d[parts[0]], parts[1:], value)
    return d

def nest_schedules( schedules:dict):

    nested_schedules = {}
    def _nest(schedule):
        set_in(nested_schedules, 
               [f'{schedule.get("space_type").replace(" ","")}-{schedule.get("schedule_type")}',schedule.get("day_type") ],
                {"data":[], "day_type":schedule.get("day_type") } )

    list( map(_nest, schedules ) )

    def _update_data(schedule):
        path = f'{schedule.get("space_type").replace(" ","")}-{schedule.get("schedule_type")}'
        nested_schedules[path][schedule.get("day_type") ]["data"].append( 
            {"from":schedule.get("from"), 
              "to":schedule.get("to"), 
               "value":schedule.get("value")}
        )                  
    
    list( map(_update_data, schedules ) )

    return nested_schedules

def update_schedule_data(schedule, data):
    for values in data.get("data"):
        hour = int(values.get("to").split(":")[0])
        min = int( values.get("to").split(":")[1] )
        schedule.addValue(openstudio.Time(0, hour, min, 0), values.get("value"))

def make_default_schedule(osm, schedule_ruleset,name,  day_schedule:dict):
    default_day = schedule_ruleset.defaultDaySchedule()
    default_day.setName( name )
    update_schedule_data(default_day, day_schedule )

def make_weekend_schedule(osm, schedule_ruleset,name, schedule_dict:dict):
    rule = openstudio.model.ScheduleRule(schedule_ruleset)
    rule.setName(name)
    rule.setApplySunday(True)
    rule.setApplySaturday(True)
    day_schedule = rule.daySchedule()
    day_schedule.setName(name)
    update_schedule_data(day_schedule, schedule_dict )

def get_schedule_handler(name, schedule):
        if name in ["weekdays", "default"]:
            return make_default_schedule
        if name in ["weekend"]:
            return make_weekend_schedule

def make_schedule_ruleset(osm, name, schedules:dict):
    """ map per like { Class5OfficeBuilding-Occupancy: {"daytype"} """

    schedule_ruleset = openstudio.model.ScheduleRuleset(osm)
    schedule_ruleset.setName( name )
    
    for name,schedule_dict in schedules.items():
        # get the handler
        handler = get_schedule_handler(name, schedule_dict)
        handler(osm, schedule_ruleset, name, schedule_dict) 

    osm.getSchedules()

def create_schedule_sets(osm, schedule_sets_dict:dict):

    def _make(data:dict):
        schedule_set = openstudio.model.DefaultScheduleSet(osm)
        schedule_set.setName( data.get("name") )
        s = osm.getScheduleRulesetByName(data.get("hours_of_operation")).get()
        schedule_set.setHoursofOperationSchedule( s )
        s = osm.getScheduleRulesetByName(data.get("number_of_people")).get()
        schedule_set.setNumberofPeopleSchedule( s )
        s = osm.getScheduleRulesetByName(data.get("lighting")).get()
        schedule_set.setLightingSchedule( s )
        s = osm.getScheduleRulesetByName(data.get("electric_equipment")).get()
        schedule_set.setElectricEquipmentSchedule( s )
        s = osm.getScheduleRulesetByName(data.get("infiltration")).get()
        schedule_set.setInfiltrationSchedule( s )

    list(map( _make, schedule_sets_dict))

class AUSLibrary(openstudio.measure.ModelMeasure):
    """A ModelMeasure."""

    def name(self):

        return "AUS Library"

    def description(self):

        return "Replace this text with an explanation of what the measure does in terms that can be understood by a general building professional audience (building owners, architects, engineers, contractors, etc.).  This description will be used to create reports aimed at convincing the owner and/or design team to implement the measure in the actual building design.  For this reason, the description may include details about how the measure would be implemented, along with explanations of qualitative benefits associated with the measure.  It is good practice to include citations in the measure if the description is taken from a known source or if specific benefits are listed."

    def modeler_description(self):

        return "Replace this text with an explanation for the energy modeler specifically.  It should explain how the measure is modeled, including any requirements about how the baseline model must be set up, major assumptions, citations of references to applicable modeling resources, etc.  The energy modeler should be able to read this description and understand what changes the measure is making to the model and why these changes are being made.  Because the Modeler Description is written for an expert audience, using common abbreviations for brevity is good practice."

    def arguments(self, model: typing.Optional[openstudio.model.Model] = None):
        """Prepares user arguments for the measure.

        Measure arguments define which -- if any -- input parameters the user may set before running the measure.
        """
        args = openstudio.measure.OSArgumentVector()

        # example_arg = openstudio.measure.OSArgument.makeStringArgument("Dummy", True)
        # example_arg.setDisplayName("Dummy Arg")
        # example_arg.setDescription("This .")
        # example_arg.setDefaultValue("sd")
        # args.append(example_arg)

        return args

    def run(
        self,
        model: openstudio.model.Model,
        runner: openstudio.measure.OSRunner,
        user_arguments: openstudio.measure.OSArgumentMap,
    ):
        """Defines what happens when the measure is run."""
        super().run(model, runner, user_arguments)  # Do **NOT** remove this line

        if not (runner.validateUserArguments(self.arguments(model), user_arguments)):
            return False

        data = read_excel()

        # Make Schedules
        nested_schedules = nest_schedules( data.get("schedules"))
        [ make_schedule_ruleset(model, name, schedules) for name, schedules in nested_schedules.items()] 
        
        # Make Schedule Sets
        create_schedule_sets(model, data.get("schedule_sets"))

        # People
        create_people_load(model, data.get("people"))

        # Lights
        create_lights_load(model, data.get("lights"))

        # electric equipment defs   (BCA and that mech std)
        create_electric_equipment_load(model, data.get("equipment"))

        # infiltration defs (mostly BCA)
        create_infiltration_objects(model, data.get("infiltration"))

        # OA defs   (AS 1668.2)
        create_outdoor_air_objects(model, data.get("outdoor_air"))

        # Make Space Types
        create_space_types(model, data.get("space_types"))

        return True


# register the measure to be used by the application
AUSLibrary().registerWithApplication()
