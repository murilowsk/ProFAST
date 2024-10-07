# uncompyle6 version 3.9.2
# Python bytecode version base 3.8.0 (3413)
# Decompiled from: Python 3.11.5 | packaged by Anaconda, Inc. | (main, Sep 11 2023, 13:26:23) [MSC v.1916 64 bit (AMD64)]
# Embedded file name: C:\Users\bkee\Documents\GitHub\ProFAST\ProFAST\ProFAST.py
# Compiled at: 2024-06-10 18:22:40
# Size of source mod 2**32: 118693 bytes
import json, math, os, shutil, time
from dataclasses import dataclass
from typing import Literal
import matplotlib.pyplot as plt
import numpy as np, numpy_financial as npf, pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import xlwings as xw
from importlib_resources import files
from matplotlib import patches as mpatches
_PLOT_TYPES = Literal[('bar', 'pie')]

@dataclass
class per_kg_stream:
    name = ""
    name: str
    usage = 0
    usage: int
    unit = ""
    unit: str
    cost = 0
    cost: int
    escalation = 0
    escalation: int
    all_feedstock_regions = 0
    all_feedstock_regions: int

    def __post_init__(self):
        if isinstance(self.usage, dict):
            self.usage = {str(k): v for k, v in self.usage.items()}
        if isinstance(self.cost, dict):
            self.cost = {str(k): v for k, v in self.cost.items()}

    def update_sales(self, sales, init, year_cols, analysis_year, regional_feedstock):
        if init:
            if isinstance(self.usage, list):
                if len(self.usage) != len(year_cols):
                    raise Exception(f"Usage list must be same length as analysis period: {len(year_cols)}")
            else:
                usage = self.usage
                if isinstance(self.usage, dict):
                    usage = [self.usage[y] if str(y) in self.usage else 0 for y in year_cols]
                else:
                    self.sales = sales * usage
                    self.multiplier = np.array([1.0] * len(analysis_year))
                    if isinstance(self.cost, str):
                        if "X" in self.cost:
                            costsplit = self.cost.split("X")
                            self.multiplier = np.array([
                             float(costsplit[0].strip())] * len(analysis_year))
                            region = costsplit[1].strip()
                        else:
                            region = self.cost
                        if region in self.all_feedstock_regions and self.name in regional_feedstock["name"].values:
                            self.cost_formatted = regional_feedstock.loc[(
                             (regional_feedstock["cost_formatted"] == region) & (regional_feedstock["name"] == self.name),
                             year_cols)].values[0]
                        else:
                            raise Exception(f"Region '{region}' or feedstock name '{self.name}' is not valid")
                    else:
                        if isinstance(self.cost, dict):
                            self.cost_formatted = [self.cost[y] if str(y) in self.cost else 0 for y in year_cols]
                        else:
                            if isinstance(self.cost, list):
                                self.cost_formatted = self.cost
                            else:
                                self.cost_formatted = self.cost * (1.0 + self.escalation) ** (analysis_year - 1)
        self.value_per_unit = self.cost_formatted * self.multiplier
        self.cash_flow = self.sales * self.value_per_unit
        return self.cash_flow


@dataclass
class feedstock(per_kg_stream):
    is_revenue = False
    is_revenue: bool


@dataclass
class coproduct(per_kg_stream):
    is_revenue = True
    is_revenue: bool


@dataclass
class fixed_cost:
    name = ""
    name: str
    unit = ""
    unit: str
    cost = 0
    cost: int
    usage = 1
    usage: int
    escalation = 0
    escalation: int

    def __post_init__(self):
        if isinstance(self.usage, dict):
            self.usage = {str(k): v for k, v in self.usage.items()}
        if isinstance(self.cost, dict):
            self.cost = {str(k): v for k, v in self.cost.items()}

    def update_sales(self, init, year_cols, analysis_year, fraction_of_year_operated):
        if init:
            if isinstance(self.cost, dict):
                self.cost_formatted = [self.cost[y] if str(y) in self.cost else 0 for y in year_cols]
            else:
                if isinstance(self.cost, list):
                    self.cost_formatted = self.cost
                else:
                    self.cost_formatted = self.cost * (1.0 + self.escalation) ** (analysis_year - 1)
            self.cash_flow = self.cost_formatted * fraction_of_year_operated
        return self.cash_flow


@dataclass
class capital_item:
    name = ""
    name: str
    cost = 0
    cost: float
    depr_type = 0
    depr_type: int
    depr_period = 0
    depr_period: int
    refurb = 0
    refurb: int

    def __post_init__(self):
        if self.depr_type.lower() not in ('macrs', 'straight line'):
            raise Exception(f"{self.depr_type} is not a valid depreciation type! Capital item was not added. Try MACRS or Straight line")
        if self.depr_type.lower() == "macrs":
            MACRS_years = [
             3, 5, 7, 10, 15, 20]
            if self.depr_period not in MACRS_years:
                closest_year = min(MACRS_years, key=(lambda x: abs(x - self.depr_period)))
                raise Exception(f"{self.depr_period} is not a valid MACRS depreciation period! Value was reset to {closest_year}")
        if self.refurb == "" or self.refurb == 0:
            self.refurb = [
             0]
        if isinstance(self.refurb, dict):
            self.refurb = {str(k): v for k, v in self.refurb.items()}

    def format_refurb(self, year_cols):
        if isinstance(self.refurb, dict):
            self.refurb_list = [self.refurb[y] if str(y) in self.refurb else 0 for y in year_cols]
        else:
            self.refurb_list = self.refurb

    def update_sales(self, max_refurb_len, analysis_length):
        self.refurb_formatted = np.pad((self.refurb_list),
          pad_width=(0, max_refurb_len - len(self.refurb)))
        self.capital_exp = np.pad((np.multiply([1.0, *self.refurb_formatted], self.cost)),
          pad_width=(
         0, max(0, analysis_length - max_refurb_len)))
        return self.capital_exp

    def depreciate(self, im, MACRS_table, fyo):
        self.refurb_depr_schedule, self.depr_schedule = depreciate(self.depr_type, self.depr_period, self.refurb, self.cost, im, MACRS_table, fyo)


@dataclass
class incentive:
    __doc__ = "\n    Class for per unit commodity incentives\n    "
    name = ""
    name: str
    value = 0
    value: int
    decay = 0
    decay: int
    sunset_years = 0
    sunset_years: int
    tax_credit = True
    tax_credit: bool

    def __post_init__(self):
        if isinstance(self.value, dict):
            self.value = {str(k): v for k, v in self.value.items()}
        if self.sunset_years < 0:
            raise Exception("Sunset years must be >= 0")

    def update_sales(self, sales, init, year_cols, analysis_year, yrs_of_operation, fyo):
        if init:
            if isinstance(self.value, dict):
                self.value_per_year = [self.value[y] if str(y) in self.value else 0 for y in year_cols]
            else:
                if isinstance(self.value, list):
                    self.value_per_year = self.value
                else:
                    incentive_escalation = self.value * (1 + -1.0 * self.decay) ** (analysis_year - 1)
                    bb = np.logical_and(yrs_of_operation > 0, yrs_of_operation <= self.sunset_years)
                    cc = np.logical_and(yrs_of_operation > self.sunset_years, yrs_of_operation < self.sunset_years + 1)
                    dd = fyo
                    inc1 = incentive_escalation * bb
                    inc2 = incentive_escalation * cc * dd * (1 - np.mod(yrs_of_operation, 1))
                    self.value_per_year = inc1 + inc2
        self.revenue = self.value_per_year * sales
        return self.revenue


class ProFAST:
    __doc__ = "\n    A class to represent a ProFAST scenario\n\n    Attributes\n    ----------\n    vals : dict\n        Dictionary of all variables\n    feedstocks : dict\n        Dictionary for any feedstocks\n    coproducts : dict\n        Dictionary for any coproducts\n    fixed_costs : dict\n        Dictionary for any fixed costs\n    capital_items : dict\n        Dictionary for any capital items\n    incentives : dict\n        Dictionary for any incentives\n\n    Methods\n    -------\n    load_json(file):\n        Import a scenario from a JSON formatted input file\n    set_params(name,value):\n        Set parameter <name> to <value>\n    load_MACRE_table():\n        Load in MACRS depreciation table from csv\n    add_capital_item(name,cost,depr_type,depr_period,refurb):\n        Add a capital item\n    add_feedstock(name,usage,unit,cost,escalation):\n        Add a feedstock (expense)\n    add_coproduct(name,usage,unit,cost,escalation):\n        Add a coproduct (revenue)\n    add_fixed_cost(name,usage,unit,cost,escalation):\n        Add a fixed cost (expense)\n    add_incentive(name,value,decay,sunset_years)\n        Add an incentive (revenue)\n    edit_capital_item(name,value):\n    edit_feedstock(name,value):\n    edit_coproduct(name,value):\n    edit_fixed_cost(name,value):\n    edit_incentive(name,value):\n    remove_capital_item(name):\n    remove_feedstock(name):\n    remove_coproduct(name):\n    remove_fixed_cost(name):\n    remove_incentive(name):\n    update_sales(sales,analysis_year,analysis_length,yrs_of_operation,fraction_of_year_operated):\n        Update all revenue streams with sale of commodities and capital expenditures\n    clear_values(class_type)\n        Clears <class_type> of all entries (i.e. feedstock)\n    plot_cashflow:\n        Plots the investor cash flow\n    plot_costs:\n        Plot cost of goods breakdown\n    plot_time_series()\n        Plot yearly values in plotly\n    plot_costs_yearly:\n        Plot cost breakdown per year\n    plot_costs_yearly2:\n        Plot cost breakdown per year in interavtive plotly\n    cash_flow(price=None)\n        Calculate net present value using commodity price of <price>\n    solve_price():\n        Solve for commodity price when net present value is zero\n    loan_calc(un_depr_cap,net_cash_by_investing,receipt_one_time_cap_incentive,all_depr,earnings_before_int_tax_depr,annul_op_incent_rev,total_revenue,total_operating_expenses,net_ppe):\n        Calculate the financial loan amounts\n    depreciate(type,period,percent,cap,qpis,ypis):\n        Depreciates an item following MACRS or Straight line\n    "

    def __init__(self, case=None):
        """
        Initialization of ProFAST class

        Parameters:
        -----------
        case=None : string
            String to denote a premade json file in the resources folder
        """
        self.vals = {}
        self.feedstocks = {}
        self.coproducts = {}
        self.fixed_costs = {}
        self.capital_items = {}
        self.incentives = {}
        defaults = {'capacity':1, 
         'long term utilization':1, 
         'demand rampup':0, 
         'analysis start year':2016, 
         'operating life':1, 
         'installation months':0, 
         'TOPC':{
          'unit price': 0.0, 
          'decay': 0.0, 
          'support utilization': 0.0, 
          'sunset years': 0}, 
         'commodity':{
          'initial price': 2.0, 
          'name': '"-"', 
          'unit': '"-"', 
          'escalation': 0.0}, 
         'annual operating incentive':{
          'value': 0.0, 
          'decay': 0.0, 
          'sunset years': 0, 
          'taxable': True}, 
         'incidental revenue':{'value':0.0, 
          'escalation':0.0}, 
         'credit card fees':0.0, 
         'sales tax':0.0, 
         'road tax':{'value':0.0, 
          'escalation':0.0}, 
         'labor':{'value':0.0, 
          'rate':0.0,  'escalation':0.0}, 
         'maintenance':{'value':0.0, 
          'escalation':0.0}, 
         'rent':{'value':0.0, 
          'escalation':0.0}, 
         'license and permit':{'value':0.0, 
          'escalation':0.0}, 
         'non depr assets':0.0, 
         'end of proj sale non depr assets':0.0, 
         'installation cost':{
          'value': 0.0, 
          'depr type': '"Straight line"', 
          'depr period': 3, 
          'depreciable': True}, 
         'one time cap inct':{
          'value': 0.0, 
          'depr type': '"MACRS"', 
          'depr period': 3, 
          'depreciable': True}, 
         'property tax and insurance':0.0, 
         'admin expense':0.0, 
         'tax loss carry forward years':0, 
         'capital gains tax rate':0.0, 
         'tax losses monetized':True, 
         'sell undepreciated cap':True, 
         'loan period if used':0, 
         'debt equity ratio of initial financing':0.0, 
         'debt interest rate':0.0, 
         'debt type':"Revolving debt", 
         'total income tax rate':0.0, 
         'cash onhand':0.0, 
         'general inflation rate':0.0, 
         'leverage after tax nominal discount rate':0.0}
        self.val_names = defaults.keys()
        self.load_MACRS_table()
        self.load_feedstock_costs()
        self.default_values = self.vals
        if case is None or case == "blank":
            for i in defaults:
                self.set_params(i, defaults[i])

        else:
            self.load_json(case)

    def load_json(self, case: str):
        """
        Overview:
        ---------
            Load a ProFAST scenario based on a json case file

        Parameters:
        -----------
            case:str - The case file name e.g. central_grid_electrolysis_PEM
        """
        self.clear_values("all")
        if os.path.isfile(case):
            f = open(case)
        else:
            if os.path.isfile(files("ProFAST.resources").joinpath(f"{case}.json")):
                f = open(files("ProFAST.resources").joinpath(f"{case}.json"))
            else:
                raise ValueError(f'File location "{case}.json" not found')
        data = json.load(f)
        vars = data["variables"]
        for i in vars:
            self.set_params(i, vars[i])
        else:
            if "feedstock" in data:
                vars = data["feedstock"]
                for i in vars:
                    self.add_feedstock(i["name"], i["usage"], i["unit"], i["cost"], i["escalation"])
                else:
                    if "coproduct" in data:
                        vars = data["coproduct"]
                        for i in vars:
                            self.add_coproduct(i["name"], i["usage"], i["unit"], i["cost"], i["escalation"])

                if "fixed cost" in data:
                    vars = data["fixed cost"]
                    for i in vars:
                        self.add_fixed_cost(i["name"], i["usage"], i["unit"], i["cost"], i["escalation"])

            elif "capital item" in data:
                vars = data["capital item"]
                for i in vars:
                    self.add_capital_item(i["name"], i["cost"], i["depr type"], i["depr period"], i["refurb"])

            if "incentive" in data:
                vars = data["incentive"]
                for i in vars:
                    self.add_incentive(i["name"], i["value"], i["decay"], i["sunset years"], i["tax credit"])

    def set_params(self, name: str, value):
        """
        Set ProFAST scenario parameters:

        Parameters:
        -----------
        name : string
            Name of the pameter to change
            Valid options are:
                capacity : float
                installation cost : {value:float, depr type:"MACRS" or "Straight line", depr period:float, depreciable:bool}
                non depr assets : float
                end of proj sale non depr assets : float
                maintenance : {value:float ,escalation:float}
                one time cap inct : {value:float, depr type:"MACRS" or "Straight line", depr period:float, depreciable:bool}
                annual operating incentive : {value:float ,decay:float, sunset years: int}
                incidental revenue : {value:float, escalation:float}
                commodity : {name:string, initial price:float, unit:string, escalation:float}
                analysis start year : int
                operating life : int
                installation months : int
                demand rampup : float
                long term utilization : float
                TOPC : {unit price:float, decay:float, support utilization:float, sunset years:int}
                credit card fees : float
                sales tax : float
                road tax : {value:float, escalation:float}
                labor : {value:float, rate:float, escalation:float}
                license and permit : {value:float, escalation:float}
                rent : {value:float, escalation:float}
                property tax and insurance : float
                admin expense : float
                total income tax rate : float
                capital gains tax rate : float
                sell undepreciated cap : bool
                tax losses monetized : bool
                tax loss carry forward years : int
                general inflation rate : float
                leverage after tax nominal discount rate : float
                debt equity ratio of initial financing : float
                debt type : "Revolving debt" or "One time loan"
                loan period if used : int
                debt interest rate : float
                cash onhand : float
        """

        def type_check(name, value, types=(
 float, np.floating, int)):
            if not isinstance(value, types):
                raise ValueError(f'Parameter: "{name}" cannot be of type {type(value)}. Type must be one of {types}')

        def positive_checkParse error at or near `COME_FROM' instruction at offset 36_0

        def check_keys(name, value, keys):
            if not value.keys() >= keys:
                raise ValueError(f'Parameter: "{name}" needs to have dict keys of {keys}')

        def check_in_list(name, value, list_vals):
            if value not in list_vals:
                raise ValueError(f'Parameter: "{name}" needs to be one of the following: {list_vals}')

        def check_range(name, value, low, high):
            if value < low or value > high:
                raise ValueError(f'Parameter: "{name}" needs to be between: {low} and {high}')

        if name not in self.val_names:
            raise ValueError(f"{name} is not a valid parameter name")
        value_type = type(value)
        if name in ('capacity', 'non depr assets', 'end of proj sale non depr assets',
                    'demand rampup', 'operating life', 'installation months', 'tax loss carry forward years',
                    'loan period if used'):
            type_check(name, value)
            positive_check(name, value, name in ('capacity', 'operating life'))
        if name in ('credit card fees', 'sales tax', 'property tax and insurance',
                    'admin expense', 'total income tax rate', 'capital gains tax rate',
                    'general inflation rate', 'leverage after tax nominal discount rate',
                    'debt equity ratio of initial financing', 'debt interest rate',
                    'cash onhand'):
            type_check(name, value)
        if name in ('long term utilization', ):
            type_check(name, value, (float, np.floating, int, dict))
            if isinstance(value, dict):
                value = {str(k): v for k, v in value.items()}
        if name in ('installation cost', 'one time cap inct'):
            type_check(name, value, dict)
            check_keys(name, value, {"value", "depr type", "depr period", "depreciable"})
            type_check("{name}-value", value["value"])
            check_in_list(f"{name}-depr type", value["depr type"], ["MACRS", "Straight line"])
            if value["depr type"] == "MACRS":
                check_in_list(f"{name}-depr period", value["depr period"], [3, 5, 7, 10, 15, 20])
            type_check(f"{name}-depr period", value["depr period"])
            positive_check(f"{name}-depr period", value["depr period"], True)
            type_check(f"{name}-depreciable", (value["depreciable"]), types=bool)
        if name in ('maintenance', 'incidental revenue', 'license and permit', 'rent',
                    'road tax', 'labor'):
            type_check(name, value, dict)
            check_keys(name, value, {"value", "escalation"})
            type_check(f"{name}-value", value["value"])
            type_check(f"{name}-escalation", value["escalation"])
            if name == "labor":
                check_keys(name, value, {"rate"})
                type_check(f"{name}-rate", value["rate"])
        if name in ('analysis start year', ):
            type_check(name, value, int)
            if name == "analysis start year":
                check_range(name, value, 1000, 4000)
        if name in ('sell undepreciated cap', 'tax losses monetized'):
            type_check(name, value, bool)
        if name == "annual operating incentive":
            type_check(name, value, dict)
            check_keys(name, value, {"value", "decay", "sunset years", "taxable"})
            for i in ('value', 'decay', 'sunset years'):
                type_check(f"{name}-{i}", value[i])
            else:
                type_check(f"{name}-taxable", value["taxable"], bool)

        if name == "TOPC":
            type_check(name, value, dict)
            check_keys(name, value, {
             "unit price", "decay", "support utilization", "sunset years"})
            for i in ('unit price', 'decay', 'support utilization', 'sunset years'):
                type_check(f"{name}-{i}", value[i])

        if name == "debt type":
            check_in_list(name, value, ["Revolving debt", "One time loan"])
        if name == "commodity":
            type_check(name, value, dict)
            check_keys(name, value, {"name", "unit", "escalation"})
            type_check(f"{name}-name", value["name"], str)
            type_check(f"{name}-unit", value["unit"], str)
            type_check(f"{name}-escalation", value["escalation"])
        self.vals[name] = value

    def load_MACRS_table(self):
        """
        Read in MACRS depreciation table
        """
        macrs_table_dict = {}
        macrs_table_dict["Recovery"] = list(range(1, 22, 1))
        macrs_table_dict["Q1_3"] = [0.5833, 0.2778, 0.1235, 0.0154] + [0] * 17
        macrs_table_dict["Q1_5"] = [0.35, 0.26, 0.156, 0.1101, 0.1101, 0.0138] + [
         0] * 15
        macrs_table_dict["Q1_7"] = [
         0.25, 
         0.2143, 
         0.1531, 
         0.1093, 
         0.0875, 
         0.0874, 
         0.0875, 
         0.0109] + [
         0] * 13
        macrs_table_dict["Q1_10"] = [
         0.175, 
         0.165, 
         0.132, 
         0.1056, 
         0.0845, 
         0.0676, 
         0.0655, 
         0.0655, 
         0.0656, 
         0.0655, 
         0.0082] + [
         0] * 10
        macrs_table_dict["Q1_15"] = [
         0.0875, 
         0.0913, 
         0.0821, 
         0.0739, 
         0.0665, 
         0.0599, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.059, 
         0.0074] + [
         0] * 5
        macrs_table_dict["Q1_20"] = [
         0.06563, 
         0.07, 
         0.06482, 
         0.05996, 
         0.05546, 
         0.0513, 
         0.04746, 
         0.04459, 
         0.04459, 
         0.04459, 
         0.04459, 
         0.0446, 
         0.04459, 
         0.0446, 
         0.04459, 
         0.0446, 
         0.04459, 
         0.0446, 
         0.04459, 
         0.0446, 
         0.00565]
        macrs_table_dict["Q2_3"] = [
         0.4167, 0.3889, 0.1414, 0.053] + [0] * 17
        macrs_table_dict["Q2_5"] = [0.25, 0.3, 0.18, 0.1137, 0.1137, 0.0426] + [0] * 15
        macrs_table_dict["Q2_7"] = [
         0.1785, 
         0.2347, 
         0.1676, 
         0.1197, 
         0.0887, 
         0.0887, 
         0.0887, 
         0.0334] + [
         0] * 13
        macrs_table_dict["Q2_10"] = [
         0.125, 
         0.175, 
         0.14, 
         0.112, 
         0.0896, 
         0.0717, 
         0.0655, 
         0.0655, 
         0.0656, 
         0.0655, 
         0.0246] + [
         0] * 10
        macrs_table_dict["Q2_15"] = [
         0.0625, 
         0.0938, 
         0.0844, 
         0.0759, 
         0.0683, 
         0.0615, 
         0.0591, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.0221] + [
         0] * 5
        macrs_table_dict["Q2_20"] = [
         0.04688, 
         0.07148, 
         0.06612, 
         0.06116, 
         0.05658, 
         0.05233, 
         0.04841, 
         0.04478, 
         0.04463, 
         0.04463, 
         0.04463, 
         0.04463, 
         0.04463, 
         0.04463, 
         0.04462, 
         0.04463, 
         0.04462, 
         0.04463, 
         0.04462, 
         0.04463, 
         0.01673]
        macrs_table_dict["Q3_3"] = [
         0.25, 0.5, 0.1667, 0.0833] + [0] * 17
        macrs_table_dict["Q3_5"] = [0.15, 0.34, 0.204, 0.1224, 0.113, 0.0706] + [0] * 15
        macrs_table_dict["Q3_7"] = [
         0.1071, 
         0.2551, 
         0.1822, 
         0.1302, 
         0.093, 
         0.0885, 
         0.0886, 
         0.0553] + [
         0] * 13
        macrs_table_dict["Q3_10"] = [
         0.075, 
         0.185, 
         0.148, 
         0.1184, 
         0.0947, 
         0.0758, 
         0.0655, 
         0.0655, 
         0.0656, 
         0.0655, 
         0.041] + [
         0] * 10
        macrs_table_dict["Q3_15"] = [
         0.0375, 
         0.0963, 
         0.0866, 
         0.078, 
         0.0702, 
         0.0631, 
         0.059, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.0369] + [
         0] * 5
        macrs_table_dict["Q3_20"] = [
         0.02813, 
         0.07289, 
         0.06742, 
         0.06237, 
         0.05769, 
         0.05336, 
         0.04936, 
         0.04566, 
         0.0446, 
         0.0446, 
         0.0446, 
         0.0446, 
         0.04461, 
         0.0446, 
         0.04461, 
         0.0446, 
         0.04461, 
         0.0446, 
         0.04461, 
         0.0446, 
         0.02788]
        macrs_table_dict["Q4_3"] = [
         0.0833, 0.6111, 0.2037, 0.1019] + [0] * 17
        macrs_table_dict["Q4_5"] = [0.05, 0.38, 0.228, 0.1368, 0.1094, 0.0958] + [
         0] * 15
        macrs_table_dict["Q4_7"] = [
         0.0357, 
         0.2755, 
         0.1968, 
         0.1406, 
         0.1004, 
         0.0873, 
         0.0873, 
         0.0764] + [
         0] * 13
        macrs_table_dict["Q4_10"] = [
         0.025, 
         0.195, 
         0.156, 
         0.1248, 
         0.0998, 
         0.0799, 
         0.0655, 
         0.0655, 
         0.0656, 
         0.0655, 
         0.0574] + [
         0] * 10
        macrs_table_dict["Q4_15"] = [
         0.0125, 
         0.0988, 
         0.0889, 
         0.08, 
         0.072, 
         0.0648, 
         0.059, 
         0.059, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.059, 
         0.0591, 
         0.059, 
         0.0517] + [
         0] * 5
        macrs_table_dict["Q4_20"] = [
         0.00938, 
         0.0743, 
         0.06872, 
         0.06357, 
         0.0588, 
         0.05439, 
         0.05031, 
         0.04654, 
         0.04458, 
         0.04458, 
         0.04458, 
         0.04458, 
         0.04458, 
         0.04458, 
         0.04458, 
         0.04458, 
         0.04458, 
         0.04459, 
         0.04458, 
         0.04459, 
         0.03901]
        self.MACRS_table = pd.DataFrame(macrs_table_dict)

    def load_feedstock_costs(self):
        """
        Load in AEO2022 regional feedstock costs
        """
        self.regional_feedstock = pd.read_csv(files("ProFAST.resources").joinpath("regional_feedstock_costs.csv"))
        self.all_feedstock_regions = self.regional_feedstock["cost_formatted"].unique()
        self.regional_feedstock_names = self.regional_feedstock["name"].unique()

    def add_capital_item(self, name, cost, depr_type, depr_period, refurb):
        """
        Overview:
        ---------
            Add a capital expenditure

        Parameters:
        -----------
            name:str - Name of the capital item
            cost:float - Cost of the capital item in start year dollars
            depr_type:str - Depreciation type, must be MACRS or Straight line
            depr_period:int - Depreciation period, if MACRS, must be 3, 5, 7, 10, 15, or 20
            refurb:list[float] - Array of refubishment fractions (e.g. [0,0.1,0,0,0.5,0.1])
        """
        self.capital_items[name] = capital_item(name=name,
          cost=cost,
          depr_type=depr_type,
          depr_period=depr_period,
          refurb=refurb)

    def add_feedstock(self, name, usage, unit, cost, escalation):
        """
        Overview:
        ---------
            Add a feedstock expense

        Parameters:
        -----------
            name:str - Name of the feedstock
            usage:float - Usage of feedstock per unit of commondity
            unit:str - Unit for feedstock quantity (e.g. kg) only used for reporting
            cost:str or list or dict - Cost of the feedstock in nominal $ per unit of feedstock
            escalation:float - Yearly escalation of feedstock price
        """
        self.feedstocks[name] = feedstock(name=name,
          usage=usage,
          unit=unit,
          cost=cost,
          escalation=escalation,
          all_feedstock_regions=(self.all_feedstock_regions))

    def add_coproduct(self, name, usage, unit, cost, escalation):
        """
        Overview:
        ---------
            Add a coproduct recenue

        Parameters:
        -----------
            name:str - Name of the feedstock
            usage:float - Usage of feedstock per unit of commondity
            unit:str - Unit for feedstock quantity (e.g. kg) only used for reporting
            cost:str or list or dict - Cost of the feedstock in nominal $ per unit of coproduct
            escalation:float - Yearly escalation of feedstock price
        """
        self.coproducts[name] = coproduct(name=name,
          usage=usage,
          unit=unit,
          cost=cost,
          escalation=escalation,
          all_feedstock_regions=(self.all_feedstock_regions))

    def add_fixed_cost(self, name, usage, unit, cost, escalation):
        """
        Overview:
        ---------
            Add a yearly fixed cost

        Parameters:
        -----------
            name:str - Name of the fixed cost
            usage:float - Usage multiplier - default to 1
            unit:str - Unit of fixed cost ($)
            cost:float - Yearly cost ($)
            escalation:float - Yearly escalation of fixed cost
        """
        self.fixed_costs[name] = fixed_cost(name=name,
          usage=usage,
          unit=unit,
          cost=cost,
          escalation=escalation)

    def add_incentive(self, name, value, decay, sunset_years, tax_credit):
        """
        Overview:
        ---------
            Add a per unit commodity incentive

        Parameters:
        -----------
            name:str - Name of the incentive
            value:float - Value of incentive ($)
            decay:float - Yearly decay of incentive (fraction), negative is escalation
            sunset_years:int - Duration of incentive (years)
            tax_credit:bool - Is incentive treated as tax credit or revenue
        """
        self.incentives[name] = incentive(name=name,
          value=value,
          decay=decay,
          sunset_years=sunset_years,
          tax_credit=tax_credit)

    def edit_feedstock(self, name: str, value: dict):
        """
        Overview:
        ---------
            Edit the values of feedstock

        Parameters:
        -----------
            name:str - Name of the feedstock to edit
            value:dict - name,value pairs for parameter to edit
        """
        for key, val in value.items():
            if key not in ('usage', 'unit', 'cost', 'escalation'):
                raise Exception(f"{key} is not a valid value to edit!")
            elif name not in self.feedstocks:
                raise Exception(f"{name} does not exist!")
            if key == "cost":
                if isinstance(val, str):
                    check_value = val
                    if "X" in check_value:
                        check_value = check_value.split("X")[1].strip()
                    if check_value not in self.all_feedstock_regions:
                        raise Exception(f"{val} is not a valid region")
            setattr(self.feedstocks[name], key, val)

    def edit_coproduct(self, name: str, value: dict):
        """
        Overview:
        ---------
            Edit the values of coproduct

        Parameters:
        -----------
            name:str - Name of the coproduct to edit
            value:dict - name,value pairs for parameter to edit
        """
        for key, val in value.items():
            if key in ('usage', 'unit', 'cost', 'escalation'):
                setattr(self.coproducts[name], key, val)
            else:
                raise Exception(f"{key} is not a valid value to edit!")

    def edit_capital_item(self, name: str, value: dict):
        """
        Overview:
        ---------
            Edit the values of capital

        Parameters:
        -----------
            name:str - Name of the capital to edit
            value:dict - name,value pairs for parameter to edit
        """
        for key, val in value.items():
            if key in ('cost', 'depr_type', 'depr_period', 'refurb'):
                setattr(self.capital_items[name], key, val)
            else:
                raise Exception(f"{key} is not a valid value to edit!")

    def edit_fixed_cost(self, name: str, value: dict):
        """
        Overview:
        ---------
            Edit the values of fixed cost

        Parameters:
        -----------
            name:str - Name of the fixed cost to edit
            value:dict - name,value pairs for parameter to edit
        """
        for key, val in value.items():
            if key in ('usage', 'unit', 'cost', 'escalation'):
                setattr(self.fixed_costs[name], key, val)
            else:
                raise Exception(f"{key} is not a valid value to edit!")

    def edit_incentive(self, name: str, value: dict):
        """
        Overview:
        ---------
            Edit the values of incentive

        Parameters:
        -----------
            name:str - Name of the incentive to edit
            value:dict - name,value pairs for parameter to edit
        """
        for key, val in value.items():
            if key in ('value', 'decay', 'sunset_years', 'tax_credit'):
                setattr(self.incentives[name], key, val)
            else:
                raise Exception(f"{key} is not a valid value to edit!")

    def remove_capital_item(self, name: str):
        """
        Overview:
        ---------
            Delete a capital item

        Parameters:
        -----------
            name:str - Name of the capital to delete
        """
        del self.capital_items[name]

    def remove_feedstock(self, name: str):
        """
        Overview:
        ---------
            Delete a feedstock item

        Parameters:
        -----------
            name:str - Name of the feedstock to delete
        """
        del self.feedstocks[name]

    def remove_coproduct(self, name: str):
        """
        Overview:
        ---------
            Delete a coproduct item

        Parameters:
        -----------
            name:str - Name of the coproduct to delete
        """
        del self.coproducts[name]

    def remove_fixed_cost(self, name: str):
        """
        Overview:
        ---------
            Delete a fixed cost item

        Parameters:
        -----------
            name:str - Name of the fixed cost to delete
        """
        del self.fixed_costs[name]

    def remove_incentive(self, name: str):
        """
        Overview:
        ---------
            Delete a incentive item

        Parameters:
        -----------
            name:str - Name of the incentive to delete
        """
        del self.incentives[name]

    def update_sales(self, analysis_length, yrs_of_operation, init):
        """'
        Overview:
        ---------
            Update the feedstock,capital,coproduct,incentive and fixed cost dataframes based on commodity sales

        Parameters:
        -----------
            analysis_length:
            yrs_of_operation:
        """
        sales = self.yearly_sales
        analysis_year = np.arange(0, analysis_length + 1)
        fraction_of_year_operated = yrs_of_operation - np.concatenate((
         [
          0], yrs_of_operation[None[:-1]]))
        year_cols = list(map(str, self.calendar_year))
        self.fs_expense = 0
        for key, fs in self.feedstocks.items():
            self.fs_expense += fs.update_sales(sales, init, year_cols, analysis_year, self.regional_feedstock)
        else:
            self.coprod_revenue = 0
            for key, cp in self.coproducts.items():
                self.coprod_revenue += cp.update_sales(sales, init, year_cols, analysis_year, self.regional_feedstock)
            else:
                self.fixed_cost_expense = 0
                for key, fc in self.fixed_costs.items():
                    self.fixed_cost_expense += fc.update_sales(init, year_cols, analysis_year, fraction_of_year_operated)
                else:
                    self.capital_exp = 0
                    max_refurb_len = 0
                    for key, ci in self.capital_items.items():
                        ci.format_refurb(year_cols)
                        max_refurb_len = max(max_refurb_len, len(ci.refurb_list))
                    else:
                        for key, ci in self.capital_items.items():
                            self.capital_exp += ci.update_sales(max_refurb_len, analysis_length)
                        else:
                            if len(self.capital_items) > 0:
                                self.capital_exp = self.capital_exp[0[:analysis_length + 1]]
                            self.incentive_revenue = 0
                            self.incentive_tax_credit = [0] * len(self.calendar_year)
                            for key, ic in self.incentives.items():
                                ic_rev = ic.update_sales(sales, init, year_cols, analysis_year, yrs_of_operation, fraction_of_year_operated)
                                if ic.tax_credit:
                                    self.incentive_tax_credit += ic_rev
                                else:
                                    self.incentive_revenue += ic_rev

    def clear_values(self, class_type):
        if class_type == "feedstocks":
            self.feedstocks = {}
        if class_type == "capital":
            self.capital_items = {}
        if class_type == "incentives":
            self.incentives = {}
        if class_type == "coproducts":
            self.coproducts = {}
        if class_type == "fixed costs":
            self.fixed_costs = {}
        if class_type == "all":
            self.feedstocks = {}
            self.capital_items = {}
            self.incentives = {}
            self.coproducts = {}
            self.fixed_costs = {}

    def cash_flow(self, price: float=None, init: float=True) -> float:
        if price == None:
            price = self.vals["commodity"]["initial price"]
        else:
            financial_year = self.vals["analysis start year"] - 1
            analysis_length = int(self.vals["operating life"] + math.ceil(self.vals["installation months"] / 12))
            analysis_year = np.arange(0, analysis_length + 1)
            self.calendar_year = np.arange(financial_year, financial_year + analysis_length + 1)
            self.Nyears = len(self.calendar_year)
            yrs_of_operation = np.minimum(self.vals["operating life"], np.maximum(0, analysis_year - self.vals["installation months"] / 12))
            self.vals["fraction of year operated"] = yrs_of_operation - np.concatenate((
             [
              0], yrs_of_operation[None[:-1]]))
            if isinstance(self.vals["long term utilization"], dict):
                avg_utilization = np.array([self.vals["long term utilization"][str(x)] if str(x) in self.vals["long term utilization"] else 0 for x in self.calendar_year])
            else:
                avg_utilization = np.minimum(self.vals["long term utilization"] / (self.vals["demand rampup"] + 1) * yrs_of_operation, self.vals["long term utilization"] * np.minimum(1 + (self.vals["operating life"] + self.vals["installation months"] / 12) - analysis_year, 1))
            daily_sales = self.vals["capacity"] * avg_utilization
            yearly_sales = daily_sales * 365
            self.yearly_sales = yearly_sales

            def get_TOPC_revenue(TOPC, analysis_year, install_months, avg_utilization, yrs_of_operation):
                unit_price = TOPC["unit price"]
                decay = TOPC["decay"]
                sup_util = TOPC["support utilization"]
                ss_years = TOPC["sunset years"]
                TOPC_price = unit_price * (1.0 - decay) ** analysis_year
                shift = math.ceil(max(install_months / 12, 1))
                TOPC_price = np.concatenate((np.zeros(shift), TOPC_price[None[:-shift]]))
                TOPC_volume = np.maximum(sup_util - avg_utilization, 0) * self.vals["capacity"] * 365 * (yrs_of_operation > 0) * (yrs_of_operation <= ss_years)
                TOPC_revenue = TOPC_price * TOPC_volume
                return (TOPC_revenue, TOPC_volume)

            def incentive_revenue(incentive, yoo):
                a = incentive["value"]
                if "sunset years" in incentive:
                    b = incentive["decay"]
                    c = incentive["sunset years"]
                    aa = (a - a * b * (np.floor(yoo) - 1)) * (1 - yoo % 1)
                    bb = (a - a * b * np.floor(yoo)) * (yoo % 1)
                    aa[np.invert((np.floor(yoo) > 0) & (np.floor(yoo) <= c))] = 0
                    bb[np.invert((np.floor(yoo) > -1) & (np.floor(yoo) + 1 <= c))] = 0
                else:
                    b = incentive["escalation"]
                    aa = a * (1 + b) ** (np.floor(yoo) - 1) * (1 - yoo % 1)
                    bb = a * (1 + b) ** np.floor(yoo) * (yoo % 1)
                    aa[np.invert(np.floor(yoo) > 0)] = 0
                    bb[np.invert(np.floor(yoo) > -1)] = 0
                return aa + bb

            def escalate(value, rate, analysis_year):
                return value * (1.0 + rate) ** (analysis_year - 1)

            if isinstance(price, dict):
                sales_price = [price[y] if str(y) in price else 0 for y in list(map(str, self.calendar_year))]
            else:
                sales_price = price * (1.0 + self.vals["commodity"]["escalation"]) ** (analysis_year - 1)
        sales_revenue = sales_price * yearly_sales
        self.update_sales(analysis_length, yrs_of_operation, init)
        TOPC_revenue, TOPC_volume = get_TOPC_revenue(self.vals["TOPC"], analysis_year, self.vals["installation months"], avg_utilization, yrs_of_operation)
        annul_op_incent_rev = incentive_revenue(self.vals["annual operating incentive"], yrs_of_operation)
        incidental_rev = incentive_revenue(self.vals["incidental revenue"], yrs_of_operation)
        credit_card_fees_array = -self.vals["credit card fees"] * sales_revenue
        sales_taxes_array = -self.vals["sales tax"] * sales_revenue
        road_taxes = -escalate(self.vals["road tax"]["value"], self.vals["road tax"]["escalation"], analysis_year) * yearly_sales
        total_revenue = sales_revenue + self.coprod_revenue + annul_op_incent_rev + self.incentive_revenue + TOPC_revenue + incidental_rev + credit_card_fees_array + sales_taxes_array + road_taxes
        labor_rate = escalate(self.vals["labor"]["rate"], self.vals["labor"]["escalation"], analysis_year)
        maintenance = escalate(self.vals["maintenance"]["value"], self.vals["maintenance"]["escalation"], analysis_year)
        rent_of_land = escalate(self.vals["rent"]["value"], self.vals["rent"]["escalation"], analysis_year)
        license_and_permit = escalate(self.vals["license and permit"]["value"], self.vals["license and permit"]["escalation"], analysis_year)
        labor_expense = self.vals["fraction of year operated"] * labor_rate * self.vals["labor"]["value"]
        maintenance_expense = self.vals["fraction of year operated"] * maintenance
        rent_expense = self.vals["fraction of year operated"] * rent_of_land
        license_and_permit_expense = self.vals["fraction of year operated"] * license_and_permit
        receipt_one_time_cap_incentive = np.pad([
         self.vals["one time cap inct"]["value"]],
          pad_width=(0, analysis_length))
        non_depr_assets_exp = np.pad([
         self.vals["non depr assets"]],
          pad_width=(0, analysis_length))
        installation_expenditure = np.pad([
         self.vals["installation cost"]["value"]],
          pad_width=(0, analysis_length))
        net_cash_by_investing = -1 * (self.capital_exp + non_depr_assets_exp + installation_expenditure)
        all_plant_prop_and_equip = np.cumsum(-1 * net_cash_by_investing)
        max_cap_years = 0
        for key, ci in self.capital_items.items():
            ci.depreciate(self.vals["installation months"], self.MACRS_table, self.vals["fraction of year operated"])
            max_cap_years = max(max_cap_years, len(ci.refurb_depr_schedule), len(ci.depr_schedule))
        else:
            _, installation_depr_sch = depreciate(self.vals["installation cost"]["depr type"], self.vals["installation cost"]["depr period"], [
             0], self.vals["installation cost"]["value"] * self.vals["installation cost"]["depreciable"], self.vals["installation months"], self.MACRS_table, self.vals["fraction of year operated"])
            _, cap_inct_depr_sch = depreciate(self.vals["one time cap inct"]["depr type"], self.vals["one time cap inct"]["depr period"], [
             0], -1 * self.vals["one time cap inct"]["value"] * (not self.vals["one time cap inct"]["depreciable"]), self.vals["installation months"], self.MACRS_table, self.vals["fraction of year operated"])
            max_depr_years = max(max_cap_years, len(installation_depr_sch), len(cap_inct_depr_sch), len(self.vals["fraction of year operated"]))
            refurb_depr_schedule_sum = np.zeros(max_depr_years)
            depr_schedule_sum = np.zeros(max_depr_years)
            for key, ci in self.capital_items.items():
                ci.refurb_depr_schedule = np.pad((ci.refurb_depr_schedule),
                  pad_width=(
                 0, max_depr_years - len(ci.refurb_depr_schedule)))
                ci.depr_schedule = np.pad((ci.depr_schedule),
                  pad_width=(0, max_depr_years - len(ci.depr_schedule)))
                refurb_depr_schedule_sum += ci.refurb_depr_schedule
                depr_schedule_sum += ci.depr_schedule
            else:
                installation_depr_sch = np.pad(installation_depr_sch,
                  pad_width=(
                 0, max_depr_years - len(installation_depr_sch)))
                cap_inct_depr_sch = np.pad(cap_inct_depr_sch,
                  pad_width=(0, max_depr_years - len(cap_inct_depr_sch)))
                all_depr = depr_schedule_sum + refurb_depr_schedule_sum + installation_depr_sch + cap_inct_depr_sch
                cum_depr = np.cumsum(all_depr)
                max_depr_len = max(len(all_depr) - np.where(np.abs(all_depr[None[None:-1]]) < 1e-06)[0][-2], self.Nyears)
                if len(self.capital_items) > 0:
                    cum_depr_cap = [
                     0.0] * len(ci.refurb_formatted)
                    for key, ci in self.capital_items.items():
                        cum_depr_cap += np.multiply(ci.cost, ci.refurb_formatted) + [
                         
                          ci.cost,
                         *np.zeros(len(ci.refurb_formatted) - 1)]

                else:
                    cum_depr_cap = [
                     0.0]
                install_depr_cap = np.multiply([
                 
                  self.vals["installation cost"]["value"], *np.zeros(len(cum_depr_cap) - 1)], self.vals["installation cost"]["depreciable"])
                one_time_inct_depr_cap = np.multiply([
                 
                  self.vals["one time cap inct"]["value"], *np.zeros(len(cum_depr_cap) - 1)], not self.vals["one time cap inct"]["depreciable"])
                cum_depr_cap += install_depr_cap - one_time_inct_depr_cap
                cum_depr_cap = np.concatenate((
                 np.cumsum(cum_depr_cap),
                 np.ones(len(cum_depr) - len(cum_depr_cap)) * np.cumsum(cum_depr_cap)[-1]))
                un_depr_cap = cum_depr_cap - cum_depr
                cum_PPE = np.cumsum(-net_cash_by_investing)
                cum_depr = np.multiply(-1, [0, *cum_depr][None[:max_depr_len]])
                net_ppe = np.pad(cum_PPE, pad_width=(0, max_depr_len - len(cum_PPE))) + cum_depr
                property_insurance = self.vals["fraction of year operated"] * net_ppe[None[:self.Nyears]] * self.vals["property tax and insurance"]
                admin_expense = (sales_revenue + self.coprod_revenue) * self.vals["admin expense"]
                total_operating_expenses = self.fs_expense + labor_expense + rent_expense + property_insurance + license_and_permit_expense + admin_expense + maintenance_expense + self.fixed_cost_expense
                earnings_before_int_tax_depr = total_revenue - total_operating_expenses
                NPV = self.loan_calc(un_depr_cap, net_cash_by_investing, receipt_one_time_cap_incentive, all_depr, earnings_before_int_tax_depr, annul_op_incent_rev, total_revenue, total_operating_expenses, net_ppe)
                comod_unit = self.vals["commodity"]["unit"]
                comod_name = self.vals["commodity"]["name"]
                all_depr_crop = np.concatenate(([0], all_depr[0[:len(self.calendar_year) - 1]]))
                cash_flow_out = {"Year": (self.calendar_year), 
                 "Cumulative cash flow": (np.cumsum(self.loan_out["investor_cash_flow"])), 
                 "Investor cash flow": (self.loan_out["investor_cash_flow"]), 
                 "Monetized tax losses": (-1 * np.minimum(self.loan_out["income_taxes_payable"], 0)), 
                 
                 "Gross margin": (np.divide((total_revenue - total_operating_expenses),
                                   total_revenue,
                                   out=(np.zeros_like(total_revenue)),
                                   where=(total_revenue != 0))), 
                 
                 "Cost of goods sold ($/year)": (total_operating_expenses - admin_expense + all_depr_crop + self.loan_out["interest_pmt"][None[:self.Nyears]]), 
                 
                 f"Cost of goods sold ($/{comod_unit})": (np.divide((total_operating_expenses - admin_expense + all_depr_crop + self.loan_out["interest_pmt"][None[:self.Nyears]]),
                                                           yearly_sales,
                                                           out=(np.zeros_like(yearly_sales)),
                                                           where=(yearly_sales > 0))), 
                 
                 "Average utilization": avg_utilization, 
                 f"{comod_name} sales ({comod_unit}/day)": daily_sales, 
                 f"Capacity covered by TOPC ({comod_unit}/day)": (TOPC_volume / 365), 
                 f"Cost of {comod_name} ($/{comod_unit})": sales_price, 
                 f"{comod_name} sales ({comod_unit}/year)": yearly_sales, 
                 f"Sales revenue of {comod_name} ($/year)": sales_revenue}
                for index, cp in self.coproducts.items():
                    cash_flow_out[f"Value of {cp.name} ($/{cp.unit})"] = cp.value_per_unit
                    cash_flow_out[f"{cp.name} sales ($/year)"] = cp.cash_flow
                else:
                    for index, fs in self.feedstocks.items():
                        cash_flow_out[f"Value of {fs.name} ($/{fs.unit})"] = fs.value_per_unit
                        cash_flow_out[f"{fs.name} expenses ($/year)"] = fs.cash_flow
                    else:
                        for index, ic in self.incentives.items():
                            cash_flow_out[f"Value of {ic.name} ($/{comod_unit})"] = ic.value_per_year
                            cash_flow_out[f"{ic.name} ($/year)"] = ic.revenue
                        else:
                            cash_flow_out.update({'Annual operating incentive ($/year)':annul_op_incent_rev, 
                             'Incidental revenue ($/year)':incidental_rev, 
                             'Credit card fees ($/year)':credit_card_fees_array, 
                             'Sales tax ($/year)':sales_taxes_array, 
                             'Road tax ($/year)':road_taxes, 
                             'Total revenue ($/year)':total_revenue, 
                             'Total feedstock/utilities cost ($/year)':self.fs_expense, 
                             'Labor ($/year)':labor_expense, 
                             'Total annual maintenance ($/year)':maintenance_expense, 
                             'Rent of land ($/year)':rent_expense, 
                             'Property insurance ($/year)':property_insurance, 
                             'Licensing & permitting ($/year)':license_and_permit_expense, 
                             'Administrative expense ($/year)':admin_expense, 
                             'TOPC revenue ($/year)':TOPC_revenue})
                            for index, fc in self.fixed_costs.items():
                                cash_flow_out[f"{fc.name} expenses ({fc.unit})"] = fc.cost_formatted
                            else:
                                cash_flow_out.update({'Total operating expense ($/year)':total_operating_expenses, 
                                 'EBITD ($/year)':earnings_before_int_tax_depr, 
                                 'Interest on outstanding debt ($/year)':self.loan_out["interest_pmt"], 
                                 'Depreciation ($/year)':all_depr_crop, 
                                 'Taxable income ($/year)':self.loan_out["taxable_income"], 
                                 'Remaining available deferred carry-forward tax losses ($/year)':[
                                  0] * (len(total_operating_expenses)), 
                                 'Income taxes payable ($/year)':self.loan_out["income_taxes_payable"], 
                                 'Income before extraordinary items ($/year)':self.loan_out["income_before_extraordinary_items"], 
                                 'Sale of non-depreciable assets ($/year)':self.loan_out["sale_of_non_depreciable_assets"], 
                                 'Net capital gains or loss ($/year)':(self.loan_out["sale_of_non_depreciable_assets"][None[:self.Nyears]]) - non_depr_assets_exp, 
                                 'Capital gains taxes payable ($/year)':self.loan_out["capital_gains_taxes_payable"], 
                                 'Net income ($/year)':self.loan_out["net_income"], 
                                 'Net annual operating cash flow ($/year)':(self.loan_out["net_income"][None[:self.Nyears]]) + all_depr_crop})
                                for index, ci in self.capital_items.items():
                                    cash_flow_out[f"Capital expenditure for {ci.name} ($/year)"] = ci.capital_exp[0[:len(self.calendar_year)]]
                                else:
                                    cash_flow_out.update({'Expenditure for non-depreciable fixed assets ($/year)':-non_depr_assets_exp, 
                                     'Capital expenditures for equipment installation ($/year)':installation_expenditure, 
                                     'Total capital expenditures ($/year)':-net_cash_by_investing, 
                                     'Incurrence of debt ($/year)':self.loan_out["inflow_of_debt"], 
                                     'Repayment of debt ($/year)':self.loan_out["repayment_of_debt"], 
                                     'Inflow of equity ($/year)':self.loan_out["inflow_of_equity"], 
                                     'Dividends paid ($/year)':self.loan_out["dividends_paid"], 
                                     'One-time capital incentive ($/year)':receipt_one_time_cap_incentive, 
                                     'Net cash for financing activities ($/year)':(self.loan_out["inflow_of_debt"][None[:self.Nyears]] + self.loan_out["repayment_of_debt"][None[:self.Nyears]] + self.loan_out["inflow_of_equity"][None[:self.Nyears]] + self.loan_out["dividends_paid"][None[:self.Nyears]]) + annul_op_incent_rev, 
                                     'Net change of cash ($/year)':self.loan_out["net_change_cash_equiv"], 
                                     'Cumulative cash ($/year)':self.loan_out["cumulative_cash"], 
                                     'Cumulative PP&E ($/year)':all_plant_prop_and_equip, 
                                     'Cumulative depreciation ($/year)':cum_depr, 
                                     'Net PP&E ($/year)':net_ppe, 
                                     'Cumulative deferred tax losses ($/year)':self.loan_out["cumulative_deferred_tax_losses"], 
                                     'Total assets ($/year)':self.loan_out["total_assets"], 
                                     'Cumulative debt ($/year)':self.loan_out["cumulative_debt"], 
                                     'Total liabilities ($/year)':self.loan_out["cumulative_debt"], 
                                     'Cumulative capital incentives equity ($/year)':self.loan_out["cumulative_equity_from_capital_incentives"], 
                                     'Cumulative investor equity ($/year)':self.loan_out["cumulative_equity_investor_contribution"], 
                                     'Retained earnings ($/year)':self.loan_out["retained_earnings"], 
                                     'Retained earnings no dividends ($/year)':self.loan_out["retained_earnings_no_dividends"], 
                                     'Total equity ($/year)':self.loan_out["total_equity"], 
                                     'Investor equity less capital incentive ($/year)':(self.loan_out["total_equity"]) - (self.loan_out["cumulative_equity_from_capital_incentives"]), 
                                     'Returns on investor equity':np.divide((self.loan_out["net_income"]),
                                       (self.loan_out["total_equity"] - self.loan_out["cumulative_equity_from_capital_incentives"]),
                                       out=(np.zeros_like(self.loan_out["net_income"])),
                                       where=(self.loan_out["total_equity"] - self.loan_out["cumulative_equity_from_capital_incentives"] != 0)), 
                                     'Debt/Equity ratio':np.divide((self.loan_out["cumulative_debt"]),
                                       (self.loan_out["total_equity"]),
                                       out=(np.zeros_like(self.loan_out["total_equity"])),
                                       where=(self.loan_out["total_equity"] != 0)), 
                                     'Returns on total equity':np.divide((self.loan_out["net_income"]),
                                       (self.loan_out["total_equity"]),
                                       out=(np.zeros_like(self.loan_out["total_equity"])),
                                       where=(self.loan_out["total_equity"] != 0)), 
                                     'Debt service coverage ratio (DSCR)':np.divide((total_revenue - cash_flow_out["Cost of goods sold ($/year)"]),
                                       total_revenue,
                                       out=(np.zeros_like(total_revenue)),
                                       where=(total_revenue > 0))})
                                    trim_len = len(cash_flow_out["Year"])
                                    self.cash_flow_out = {i: j[None[:trim_len]] if (isinstance(j, (list, np.ndarray)) and len(j) > trim_len) else j for i, j in cash_flow_out.items()}
                                    res = np.roots(self.loan_out["investor_cash_flow"][None[None:-1]])
                                    mask = (res.imag == 0) & (res.real > 0)
                                    res = res[mask].real
                                    self.irr = 1 / res - 1
                                    self.profit_index = -1
                                    if self.loan_out["investor_cash_flow"][0] != 0:
                                        self.profit_index = -1 * npf.npv(self.vals["general inflation rate"], [
                                         0, *self.loan_out["investor_cash_flow"][1[:None]]]) / self.loan_out["investor_cash_flow"][0]
                                    else:
                                        self.cum_cash_flow = cash_flow_out["Cumulative cash flow"]
                                        self.LCO = npf.npv(self.vals["general inflation rate"], sales_revenue / (sum(self.yearly_sales) / (1 + self.vals["general inflation rate"])))
                                        positive_flow = np.where(self.cum_cash_flow > 0)[0]
                                        positive_EBITD = np.where(earnings_before_int_tax_depr > 0)[0]
                                        if len(positive_flow) > 0:
                                            self.first_year_positive = positive_flow[0]
                                        else:
                                            self.first_year_positive = -1
                                        if len(positive_EBITD) > 0:
                                            self.first_year_positive_EBITD = positive_EBITD[0]
                                        else:
                                            self.first_year_positive_EBITD = -1
                                    return NPV

    def solve_price(self, guess_value=1):
        t1 = time.time()
        iters = 20
        P = np.zeros(iters)
        P[0] = guess_value
        NPV = np.zeros(iters)
        price = P[0]
        NPV[0] = self.cash_flow(price, True)
        P[1] = guess_value - 1 if NPV[0] > 0 else guess_value + 1
        for i in range(1, iters - 1):
            price = P[i]
            NPV[i] = self.cash_flow(price, False)
            if abs(NPV[i]) < 1e-05:
                break
            if abs(P[i] - P[i - 1]) < 0.0001:
                break
            slope = (NPV[i] - NPV[i - 1]) / (P[i] - P[i - 1])
            intercept = NPV[i] - slope * P[i]
            P[i + 1] = -intercept / slope
        else:
            timing = time.time() - t1
            return_vals = {'NPV':NPV[i], 
             'price':price, 
             'irr':self.irr, 
             'profit index':self.profit_index, 
             'investor payback period':self.first_year_positive, 
             'first year positive EBITD':self.first_year_positive_EBITD, 
             'timing':timing, 
             'lco':self.LCO}
            return return_vals

    def loan_calc(self, un_depr_cap, net_cash_by_investing, receipt_one_time_cap_incentive, all_depr, earnings_before_int_tax_depr, annul_op_incent_rev, total_revenue, total_operating_expenses, net_ppe):
        max_len = len(net_ppe)
        install_yrs = int(np.ceil(self.vals["installation months"] / 12))
        final_year = self.Nyears - 1
        un_depr_cap = un_depr_cap[None[:max_len]]
        all_depr = all_depr[None[:max_len]]
        receipt_one_time_cap_incentive = np.pad(receipt_one_time_cap_incentive,
          pad_width=(
         0, max_len - len(receipt_one_time_cap_incentive)))
        earnings_before_int_tax_depr = np.pad(earnings_before_int_tax_depr,
          pad_width=(
         0, max_len - len(earnings_before_int_tax_depr)))
        annul_op_incent_rev = np.pad(annul_op_incent_rev,
          pad_width=(0, max_len - len(annul_op_incent_rev)))
        incentive_tax_credit = np.pad((self.incentive_tax_credit),
          pad_width=(
         0, max_len - len(self.incentive_tax_credit)))
        total_revenue = np.pad(total_revenue,
          pad_width=(0, max_len - len(total_revenue)))
        total_operating_expenses = np.pad(total_operating_expenses,
          pad_width=(
         0, max_len - len(total_operating_expenses)))
        net_cash_by_investing = np.pad(net_cash_by_investing,
          pad_width=(0, max_len - len(net_cash_by_investing)))
        fyo = np.pad((self.vals["fraction of year operated"]),
          pad_width=(
         0, max_len - len(self.vals["fraction of year operated"])))
        inflow_of_debt, repayment_of_debt, inflow_of_equity, dividends_paid, interest_pmt, cumulative_debt, taxable_income, net_income, net_cash, cumulative_cash, cumulative_tax_loss_carryforward, net_cash_in_financing, net_change_cash_equiv, income_taxes_payable, income_before_extraordinary_items, sale_of_non_depreciable_assets, less_initial_cost = (np.zeros(max_len) for i in range(17))
        dt = self.vals["tax loss carry forward years"]
        deferments = np.zeros((dt + 1, max_len + 1))
        loan_repayments = np.zeros([
         self.Nyears, self.Nyears + self.vals["loan period if used"]])
        loan_interest = np.zeros([
         self.Nyears, self.Nyears + self.vals["loan period if used"]])
        sale_of_non_depreciable_assets[final_year] = self.vals["end of proj sale non depr assets"]
        less_initial_cost[final_year] = -1 * self.vals["non depr assets"]
        net_gain_or_loss_sale_non_depreciable_assets = sale_of_non_depreciable_assets + less_initial_cost
        capital_gains_taxes_payable = net_gain_or_loss_sale_non_depreciable_assets * self.vals["capital gains tax rate"]
        if not self.vals["tax losses monetized"]:
            capital_gains_taxes_payable = np.maximum(capital_gains_taxes_payable, 0)
        sale_residual_undepreciated_assets = np.zeros(max_len)
        loss_residual_undepreciated_assets = np.zeros(max_len)
        sale_residual_undepreciated_assets[final_year] = un_depr_cap[final_year - 1] * self.vals["sell undepreciated cap"]
        loss_residual_undepreciated_assets[final_year] = -un_depr_cap[final_year - 1] * (not self.vals["sell undepreciated cap"])
        extraordinary_items_after_tax = sale_of_non_depreciable_assets - capital_gains_taxes_payable + sale_residual_undepreciated_assets + loss_residual_undepreciated_assets
        inflow_of_equity[0] = -(net_cash_by_investing[0] + receipt_one_time_cap_incentive[0]) / (1 + self.vals["debt equity ratio of initial financing"])
        inflow_of_debt[0] = -net_cash_by_investing[0] - inflow_of_equity[0] - receipt_one_time_cap_incentive[0]
        if self.vals["debt type"] == "One time loan":
            loan_period = self.vals["loan period if used"]
            rate = self.vals["debt interest rate"] / 12
            per = np.arange(1, 12 * loan_period + 1).reshape(loan_period, 12)
            nper = loan_period * 12
            pv = inflow_of_debt[0]
            loan_repayments[(0, 1[:loan_period + 1])] = npf.ppmt(rate, per, nper, pv).sum(axis=1)
            loan_interest[(0, 1[:loan_period + 1])] = npf.ipmt(rate, per, nper, pv).sum(axis=1)
        depreciation_expense = np.concatenate(([0], all_depr))
        cumulative_debt[0] = inflow_of_debt[0]
        net_cash_in_financing[0] = inflow_of_debt[0] + inflow_of_equity[0] + self.vals["one time cap inct"]["value"]
        for i in range(0, max_len):
            if i < self.Nyears:
                if self.vals["debt type"] == "Revolving debt":
                    interest_pmt[i] = self.vals["debt interest rate"] * cumulative_debt[i - 1] * (1 - (self.Nyears == i + 1) * (1 - fyo[i]))
                else:
                    if self.vals["debt type"] == "One time loan":
                        interest_pmt[i] = -loan_interest.sum(axis=0)[i]
            taxable_income[i] = earnings_before_int_tax_depr[i] - interest_pmt[i] - depreciation_expense[i] - (not self.vals["annual operating incentive"]["taxable"]) * annul_op_incent_rev[i]
            deferments[(0, i)] = taxable_income[i] * self.vals["total income tax rate"] - incentive_tax_credit[i]
            ND = len(deferments)
            y = np.flip(np.arange(1, ND))

        if self.vals["tax losses monetized"]:
            deferments[(1[:None], i)] = 0
        else:
            deferments[(dt, i)] = deferments[(dt, i)] + min(deferments[(dt - 1, i - 1)] + max(deferments[(0, i)], 0), 0)
            for Y in range(2, ND):
                yy = y[Y - 1]
                this_year = deferments[(None[:None], i)]
                last_year = deferments[(None[:None], i - 1)]
                if sum(deferments[(yy[:None], i)]) == 0:
                    deferments[(yy, i)] = min(sum(last_year[yy[:None]]) + min(last_year[yy - 1], 0) + max(this_year[0], 0), 0)
                else:
                    deferments[(yy, i)] = min(last_year[yy - 1], 0)
                if not self.vals["tax losses monetized"]:
                    if dt > 0:
                        income_taxes_payable[i] = float(max(0, max(0, deferments[(0, i)]) + sum(np.minimum(0, deferments[(0[:dt], i - 1)]))))
                    else:
                        income_taxes_payable[i] = np.maximum(deferments[(0, 1[:None])], np.zeros(self.Nyears))
                else:
                    income_taxes_payable[i] = deferments[(0, i)]
                income_before_extraordinary_items[i] = total_revenue[i] - total_operating_expenses[i] - interest_pmt[i] - depreciation_expense[i] - income_taxes_payable[i]
                net_income[i] = income_before_extraordinary_items[i] + extraordinary_items_after_tax[i]
                net_cash[i] = net_income[i] + depreciation_expense[i]
                cumulative_cash[i] = (total_operating_expenses[i] + interest_pmt[i] + income_taxes_payable[i]) / 12 * (i < self.Nyears - 1) * self.vals["cash onhand"]
                net_change_cash_equiv[i] = cumulative_cash[i] - cumulative_cash[i - 1]
                net_cash_in_financing[i] = net_change_cash_equiv[i] - net_cash_by_investing[i] - net_cash[i]
                if i == self.Nyears - 1:
                    repayment_of_debt[i] = -cumulative_debt[i - 1]
                else:
                    if self.vals["debt type"] == "One time loan":
                        repayment_of_debt[i] = loan_repayments.sum(axis=0)[i]
                    elif i > 0:
                        if net_cash_by_investing[i] < 0:
                            inflow_of_debt[i] = max(net_cash_in_financing[i] - repayment_of_debt[i] - receipt_one_time_cap_incentive[i], 0) / (1 + self.vals["debt equity ratio of initial financing"]) * self.vals["debt equity ratio of initial financing"]
                    if self.vals["debt type"] == "One time loan":
                        loan_period = self.vals["loan period if used"]
                        rate = self.vals["debt interest rate"] / 12
                        per = np.arange(1, 12 * loan_period + 1).reshape(loan_period, 12)
                        nper = loan_period * 12
                        pv = inflow_of_debt[i]
                        loan_repayments[(i, (i + 1)[:loan_period + i + 1])] = npf.ppmt(rate, per, nper, pv).sum(axis=1)
                        loan_interest[(i, (i + 1)[:loan_period + i + 1])] = npf.ipmt(rate, per, nper, pv).sum(axis=1)
                    if i > 0:
                        inflow_of_equity[i] = max(net_cash_in_financing[i] - repayment_of_debt[i] - receipt_one_time_cap_incentive[i], 0)
                        if net_cash_by_investing[i] != 0:
                            inflow_of_equity[i] = inflow_of_equity[i] / (1 + self.vals["debt equity ratio of initial financing"])
                    cumulative_debt[i] = cumulative_debt[i - 1] + inflow_of_debt[i] + repayment_of_debt[i]
            else:
                dividends_paid = np.clip((net_cash_in_financing - inflow_of_debt - repayment_of_debt - receipt_one_time_cap_incentive),
                  a_min=None,
                  a_max=0)
                dividends_paid[0] = 0
                total_liabilities = cumulative_debt
                cumulative_tax_loss_carryforward = np.concatenate((
                 [
                  0], -deferments[(1[:dt + 1], 1[:None])].sum(axis=0)))[None[:max_len]]
                total_assets = net_ppe + cumulative_cash + cumulative_tax_loss_carryforward
                cumulative_equity_from_capital_incentives = np.cumsum(receipt_one_time_cap_incentive)
                cumulative_equity_investor_contribution = np.cumsum(inflow_of_equity)
                retained_earnings = np.cumsum(net_income) + np.cumsum(dividends_paid)
                cumulative_deferred_tax_losses = cumulative_tax_loss_carryforward
                total_equity = cumulative_equity_from_capital_incentives + cumulative_equity_investor_contribution + retained_earnings + cumulative_deferred_tax_losses
                AmLmEcheck = total_assets - total_equity - total_liabilities
                investor_cash_flow = -(inflow_of_equity + dividends_paid)
                NPV = npf.npv(self.vals["leverage after tax nominal discount rate"], investor_cash_flow)
                monetized_tax_losses = np.minimum(income_taxes_payable, 0) - np.minimum(capital_gains_taxes_payable, 0)
                self.loan_out = {'sale_of_non_depreciable_assets':sale_residual_undepreciated_assets + sale_of_non_depreciable_assets, 
                 'net_change_cash_equiv':net_change_cash_equiv, 
                 'inflow_of_equity':inflow_of_equity, 
                 'inflow_of_debt':inflow_of_debt, 
                 'dividends_paid':dividends_paid, 
                 'income_taxes_payable':income_taxes_payable, 
                 'repayment_of_debt':repayment_of_debt, 
                 'interest_pmt':interest_pmt, 
                 'capital_gains_taxes_payable':capital_gains_taxes_payable, 
                 'monetized_tax_losses':monetized_tax_losses, 
                 'depreciation_expense':depreciation_expense, 
                 'taxable_income':taxable_income, 
                 'income_before_extraordinary_items':income_before_extraordinary_items, 
                 'net_income':net_income, 
                 'cumulative_cash':cumulative_cash, 
                 'cumulative_deferred_tax_losses':cumulative_deferred_tax_losses, 
                 'total_assets':total_assets, 
                 'cumulative_debt':cumulative_debt, 
                 'cumulative_equity_from_capital_incentives':cumulative_equity_from_capital_incentives, 
                 'cumulative_equity_investor_contribution':cumulative_equity_investor_contribution, 
                 'retained_earnings':retained_earnings, 
                 'retained_earnings_no_dividends':(np.cumsum)(net_income), 
                 'total_equity':total_equity, 
                 'investor_cash_flow':investor_cash_flow}
                return NPV

    def get_summary_vals(self):
        summary_vals = {'Type':[],  'Name':[],  'Amount':[]}
        comod_name = self.vals["commodity"]["name"]
        names = [
         f'{self.vals["commodity"]["name"]} sales',
         "Take or pay revenue",
         "Incidental revenue",
         "Sale of non-depreciable assets",
         "Cash on hand recovery"]
        amount = [
         self.cash_flow_out[f"Sales revenue of {comod_name} ($/year)"],
         self.cash_flow_out["TOPC revenue ($/year)"],
         self.cash_flow_out["Incidental revenue ($/year)"],
         self.cash_flow_out["Sale of non-depreciable assets ($/year)"],
         np.minimum(self.cash_flow_out["Net change of cash ($/year)"], 0)]
        for i, cp in self.coproducts.items():
            names.append(cp.name)
            amount.append(cp.cash_flow)
        else:
            summary_vals["Name"].extend(names)
            summary_vals["Type"].extend(["Operating Revenue"] * len(names))
            summary_vals["Amount"].extend(amount)
            names = [
             'Property insurance', 
             'Road tax', 
             'Credit card fees', 
             'Sales tax', 
             'Installation cost', 
             'Total annual maintenance', 
             'Cash on hand reserve', 
             'Non-depreciable assets', 
             'Labor', 
             'Administrative expenses', 
             'Rent of land', 
             'Licensing and Permitting']
            amount = [
             self.cash_flow_out["Property insurance ($/year)"],
             self.cash_flow_out["Road tax ($/year)"],
             self.cash_flow_out["Credit card fees ($/year)"],
             self.cash_flow_out["Sales tax ($/year)"],
             self.cash_flow_out["Capital expenditures for equipment installation ($/year)"],
             self.cash_flow_out["Total annual maintenance ($/year)"],
             np.maximum(self.loan_out["net_change_cash_equiv"], 0),
             self.cash_flow_out["Expenditure for non-depreciable fixed assets ($/year)"],
             self.cash_flow_out["Labor ($/year)"],
             self.cash_flow_out["Administrative expense ($/year)"],
             self.cash_flow_out["Rent of land ($/year)"],
             self.cash_flow_out["Licensing & permitting ($/year)"]]
            for i, fs in self.feedstocks.items():
                names.append(fs.name)
                amount.append(fs.cash_flow)
            else:
                for i, fc in self.fixed_costs.items():
                    names.append(fc.name)
                    amount.append(fc.cash_flow)
                else:
                    summary_vals["Name"].extend(names)
                    summary_vals["Type"].extend(["Operating Expenses"] * len(names))
                    summary_vals["Amount"].extend(amount)
                    names = [
                     'Inflow of equity', 
                     'Inflow of debt', 
                     'Monetized tax losses', 
                     'One time capital incentive', 
                     'Annual operating incentives']
                    amount = [
                     self.loan_out["inflow_of_equity"],
                     self.loan_out["inflow_of_debt"],
                     self.loan_out["monetized_tax_losses"],
                     self.cash_flow_out["One-time capital incentive ($/year)"],
                     self.cash_flow_out["Annual operating incentive ($/year)"]]
                    for i, ic in self.incentives.items():
                        if ic.tax_credit:
                            pass
                        else:
                            names.append(ic.name)
                            amount.append(ic.revenue)
                    else:
                        summary_vals["Name"].extend(names)
                        summary_vals["Type"].extend(["Financing cash inflow"] * len(names))
                        summary_vals["Amount"].extend(amount)
                        names = [
                         'Dividends paid', 
                         'Income taxes payable', 
                         'Repayment of debt', 
                         'Interest expense', 
                         'Capital gains taxes payable']
                        amount = [
                         self.loan_out["dividends_paid"],
                         np.maximum(self.loan_out["income_taxes_payable"], 0),
                         self.loan_out["repayment_of_debt"],
                         self.loan_out["interest_pmt"],
                         self.loan_out["capital_gains_taxes_payable"]]
                        for i, ci in self.capital_items.items():
                            names.append(ci.name)
                            amount.append(ci.capital_exp)
                        else:
                            summary_vals["Name"].extend(names)
                            summary_vals["Type"].extend(["Financing cash outflow"] * len(names))
                            summary_vals["Amount"].extend(amount)
                            summary_vals["Name"].append("Depreciation expense")
                            summary_vals["Type"].append("NA")
                            summary_vals["Amount"].append(self.loan_out["depreciation_expense"])
                            return summary_vals

    def get_cost_breakdown(self, per_volume=True):
        summary_vals = self.get_summary_vals()
        rate = self.vals["general inflation rate"]
        volume = 1
        if per_volume:
            volume = sum(self.yearly_sales) / (1 + rate)
        summary_vals = pd.DataFrame(summary_vals)
        summary_vals["NPV"] = summary_vals["Amount"].apply(lambda x: abs(npf.npv(rate, x / volume)))
        summary_vals = summary_vals.sort_values(by=["Type", "NPV"], ascending=False)
        inflow = summary_vals.loc[summary_vals["Type"].isin(["Operating Revenue", "Financing cash inflow"])].sort_values("NPV",
          ascending=True)
        outflow = summary_vals.loc[summary_vals["Type"].isin(["Operating Expenses", "Financing cash outflow"])].sort_values("NPV",
          ascending=True)
        all_flow = pd.concat([outflow, inflow]).reset_index(drop=True)
        return all_flow

    def plot_costs(self, fileout='', show_plot=True):
        all_flow = self.get_cost_breakdown()
        all_flow = all_flow.loc[all_flow["NPV"] != 0]
        colors = all_flow.loc[(None[:None], ["Name", "Type"])]
        colors["Color"] = ""
        color_vals = {
         'Operating Revenue': '"#2626eb"', 
         'Financing cash inflow': '"#bdbdf2"', 
         'Operating Expenses': '"#f59342"', 
         'Financing cash outflow': '"#f0c099"'}
        colors["Color"] = colors["Type"].map(color_vals)
        ax = all_flow.plot.barh(x="Name",
          y="NPV",
          figsize=(8, 9),
          color=(colors["Color"].values))
        plt.subplots_adjust(left=0.5, bottom=0.05, right=0.9, top=0.95)
        plt.ylabel("")
        plt.title("Real levelized value breakdown of " + self.vals["commodity"]["name"] + " ($/" + self.vals["commodity"]["unit"] + ")")
        plt.xlim(right=(max(all_flow["NPV"]) * 1.25))
        handles = [mpatches.Patch(color=(color_vals[i])) for i in color_vals]
        labels = [f"{i}" for i in color_vals]
        plt.legend(handles, labels, loc="lower right")
        ax.bar_label((ax.containers[0]), fmt="%0.2f", fontsize=6)
        if fileout != "":
            plt.savefig(fileout, transparent=True)
        if show_plot:
            plt.show()
        return all_flow

    def plot_cashflow(self, scale='M', fileout='', show_plot=True):
        if scale == "M":
            scale_value = 1e-06
        else:
            if scale == "B":
                scale_value = 1e-09
            else:
                if scale == "":
                    scale_value = 1
        plot_data = self.loan_out["investor_cash_flow"][0[:len(self.calendar_year)]]
        fig, ax = plt.subplots(figsize=(9, 4))
        bar_x = self.calendar_year
        bar_height = plot_data * scale_value
        bar_plot = plt.bar((bar_x[plot_data >= 0]),
          (bar_height[plot_data >= 0]), color="blue")
        bar_plot2 = plt.bar((bar_x[plot_data < 0]),
          (bar_height[plot_data < 0]), color="red")
        plt.xticks(rotation=90, labels=(self.calendar_year), ticks=(self.calendar_year))
        plt.ylim(top=(max(bar_height) * 3))
        ax.bar_label(bar_plot, rotation=90, padding=8, fmt="%i")
        ax.bar_label(bar_plot2, rotation=90, label_type="center", fmt="%i")
        ax.set(ylabel=("$ (%s US)" % scale), xlabel="Year")
        plt.subplots_adjust(left=0.05, bottom=0.2, right=0.95, top=0.9)
        name = "Cash Flow"
        plt.title(name, loc="center")
        plt.tight_layout()
        if fileout != "":
            plt.savefig(fileout, transparent=True)
        if show_plot:
            plt.show()

    def plot_capital_expenses(self, scale='M', fileout='', show_plot=True, plot_type: _PLOT_TYPES='pie'):
        if scale == "M":
            scale_value = 1e-06
        else:
            if scale == "B":
                scale_value = 1e-09
            else:
                if scale == "":
                    scale_value = 1
                else:
                    name = []
                    cost = []
                    for key, ci in self.capital_items.items():
                        name.append(ci.name)
                        cost.append(ci.cost)
                    else:
                        plot_df = pd.DataFrame({'name':name,  'cost':cost})
                        plot_df = plot_df.rename(columns={"name": "Name"})
                        series = pd.Series((plot_df["cost"].values * scale_value),
                          index=(plot_df["Name"]), name="")
                        if plot_type == "bar":
                            ax = pd.DataFrame(series).T.plot.bar(stacked=True, legend=False)
                            for i in range(0, len(ax.containers)):
                                c = ax.containers[i]
                                print(c)
                                print(plot_df["Name"][i])
                                labels = [str(plot_df["Name"][i]) + str("\n") + str(round((v.get_height()), ndigits=2)) if v.get_height() > 0 else "" for v in c]
                                ax.bar_label(c, labels=labels, label_type="center")
                            else:
                                ax.set(ylabel=("$ (%s US)" % scale), xlabel=None)

                        else:
                            if plot_type == "pie":
                                plot_df.set_index("Name", inplace=True)
                                print(plot_df)
                                ax = plot_df.plot.pie(y="cost", figsize=(12, 8), legend=False)
                                ax.set_ylabel("")
                            else:
                                raise ValueError(f"plot_type must be one of {_PLOT_TYPES}")

                plt.title("Capital Expenditures by System", loc="center")
                plt.tight_layout()
                if fileout != "":
                    plt.savefig(fileout, transparent=True)
                if show_plot:
                    plt.show()

    def plot_time_series(self, fileout='', show_plot=True):
        """
        Overview:
        ---------
            This function produces a plotly graph of all the time series data (e.g., Cumulative cash flow vs year)

        Parameters:
        -----------

        Returns:
        --------
        """
        df = pd.DataFrame(self.cash_flow_out)
        y_vars = df.loc[(None[:None], df.columns != "Year")].columns
        fig = go.Figure()
        first_val = True
        buttons = []
        for y in y_vars:
            fig.add_trace(go.Bar(x=(df["Year"]),
              y=(df[y]),
              visible=first_val,
              marker_color=(np.where(df[y] < 0, "red", "blue"))))
            first_val = False
            buttons.append(dict(method="update",
              label=y,
              args=[
             {"visible": (y_vars.isin([y]))}, {"y": (df[y])}]))
        else:
            updatemenu = [
             {'buttons':buttons, 
              'direction':"down",  'showactive':True}]
            fig.update_layout(showlegend=False, updatemenus=updatemenu)
            if fileout != "":
                plt.savefig(fileout, transparent=True)
            if show_plot:
                fig.show()

    def plot_costs_yearly(self, per_kg=True, scale='M', remove_zeros=False, remove_depreciation=False, fileout='', show_plot=True):
        rate = self.vals["general inflation rate"]
        volume = self.yearly_sales
        summary_vals = pd.DataFrame(self.get_summary_vals())
        feedstock_names = list(self.feedstocks.keys())
        fixed_costs_names = list(self.fixed_costs.keys())
        other_names = [
         'Labor', 
         'Total annual maintenance', 
         'Rent of land', 
         'Property insurance', 
         'Licensing and Permitting', 
         'Interest expense', 
         'Depreciation expense']
        names = feedstock_names + fixed_costs_names + other_names
        summary_vals = summary_vals.loc[summary_vals["Name"].isin(names)]
        if scale == "M":
            scale_value = 1e-06
        else:
            if scale == "B":
                scale_value = 1e-09
            else:
                if scale == "":
                    scale_value = 1
                else:
                    for i in np.arange(len(self.calendar_year)):
                        if per_kg:
                            summary_vals[str(self.calendar_year[i])] = summary_vals["Amount"].apply(lambda x: x[i] / volume[i])
                        else:
                            summary_vals[str(self.calendar_year[i])] = summary_vals["Amount"].apply(lambda x: x[i] * scale_value)
                    else:
                        summary_vals = summary_vals.drop(columns=["Type", "Amount"])
                        summary_vals = summary_vals.set_index("Name")
                        summary_vals = summary_vals.fillna(0)
                        summary_vals.replace([np.inf, -np.inf], 0, inplace=True)
                        summary_vals = summary_vals.T
                        if remove_zeros:
                            summary_vals = summary_vals.loc[(None[:None], (summary_vals != 0).any(axis=0))]
                        if remove_depreciation:
                            summary_vals = summary_vals.drop(columns=["Depreciation expense"])
                        ax = summary_vals.plot.bar(stacked=True, figsize=(9, 6))
                        handles, labels = ax.get_legend_handles_labels()
                        plt.legend((handles[None[None:-1]]), (labels[None[None:-1]]), loc="best", prop={"size": 6})
                        plt.title("Cost breakdown")
                        plt.xlabel("Year")
                        if per_kg:
                            plt.ylabel("$/" + self.vals["commodity"]["unit"])
                        else:
                            plt.ylabel("$ (%s US)" % scale)

                if fileout != "":
                    plt.savefig(fileout, transparent=True)
                if show_plot:
                    plt.show()

    def plot_costs_yearly2(self, per_kg=True, scale='M', remove_zeros=False, remove_depreciation=False, fileout='', show_plot=True):
        rate = self.vals["general inflation rate"]
        volume = self.yearly_sales
        summary_vals = pd.DataFrame(self.get_summary_vals())
        feedstock_names = list(self.feedstocks.keys())
        fixed_costs_names = list(self.feedstocks.keys())
        other_names = [
         'Labor', 
         'Total annual maintenance', 
         'Rent of land', 
         'Property insurance', 
         'Licensing and Permitting', 
         'Interest expense', 
         'Depreciation expense']
        cost_of_goods_sold = feedstock_names + fixed_costs_names + other_names
        summary_vals = summary_vals.loc[summary_vals["Name"].isin(cost_of_goods_sold)]
        if scale == "M":
            scale_value = 1e-06
        else:
            if scale == "B":
                scale_value = 1e-09
            else:
                if scale == "":
                    scale_value = 1
                else:
                    for i in np.arange(len(self.calendar_year)):
                        if per_kg:
                            summary_vals[str(self.calendar_year[i])] = summary_vals["Amount"].apply(lambda x: x[i] / volume[i])
                        else:
                            summary_vals[str(self.calendar_year[i])] = summary_vals["Amount"].apply(lambda x: x[i] * scale_value)
                    else:
                        summary_vals = summary_vals.drop(columns=["Type", "Amount"])
                        summary_vals = summary_vals.set_index("Name")
                        summary_vals = summary_vals.fillna(0)
                        summary_vals.replace([np.inf, -np.inf], 0, inplace=True)
                        summary_vals = summary_vals.T
                        summary_vals = summary_vals.reset_index()
                        summary_vals = summary_vals.rename(columns={"index": "Year"})
                        nTotal = summary_vals[cost_of_goods_sold].sum(axis=1)
                        if remove_zeros:
                            summary_vals = summary_vals.loc[(None[:None], (summary_vals != 0).any(axis=0))]
                        if remove_depreciation:
                            summary_vals = summary_vals.drop(columns=["Depreciation expense"])
                        y_vars = summary_vals.loc[(None[:None], summary_vals.columns != "Year")].columns
                        if per_kg:
                            labels = {"value": f'Nominal $/{self.vals["commodity"]["unit"]} of {self.vals["commodity"]["name"]}'}
                        else:
                            labels = {"value": ("$ (%s US)" % scale)}

                fig = px.bar(summary_vals,
                  x="Year",
                  y=y_vars,
                  labels=labels,
                  color_discrete_sequence=(px.colors.qualitative.Alphabet))
                button_cogs = dict(label="Total Expenses",
                  method="update",
                  args=[
                 {'visible':(y_vars.isin)(cost_of_goods_sold), 
                  'title':"All", 
                  'showlegend':True}])

                def create_layout_button(column):
                    return dict(label=column,
                      method="update",
                      args=[
                     {'visible':(y_vars.isin)([column]), 
                      'title':column, 
                      'showlegend':True}])

                fig.update_layout(legend={"traceorder": "reversed"})
                fig.update_layout(updatemenus=[
                 go.layout.Updatemenu(active=0,
                   buttons=([
                  button_cogs] + list(y_vars.map(lambda column: create_layout_button(column)))))])
                if fileout != "":
                    fig.write_html(fileout)
                if show_plot:
                    fig.show()

    def export_to_H2FAST(self, filename):

        def set_capital(sheet, refurb_sheet, name, val, depr_type, depr_period, refurb, index):
            row = 56 + index
            if index > 9:
                return
            sheet[f"C{row}"].value = name
            sheet[f"D{row}"].value = val
            sheet[f"E{row}"].value = depr_type
            sheet[f"F{row}"].value = depr_period
            row = 72 + index
            cell_start = f"C{row}"
            rlen = len(refurb)
            cell_end = f"{chr(67 + rlen)}{row}"
            if rlen > 23:
                cell_end = f"A{chr(65 + rlen - 24)}{row}"
            refurb_sheet[f"{cell_start}:{cell_end}"].value = refurb

        def set_feedstock(sheet, name, usage, cost, escalation, unit, index):
            row = 72 + index
            row2 = 97 + index * 2
            if index > 14:
                return
            sheet[f"E{row}"].value = name
            sheet[f"D{row}"].value = usage
            sheet[f"D{row2}"].value = cost
            sheet[f"F{row}"].value = unit
            sheet[f"D{row2 + 1}"].value = escalation

        def set_coprod(sheet, name, usage, cost, escalation, unit, index):
            row = 89 + index
            row2 = 129 + index * 2
            if index > 5:
                return
            sheet[f"E{row}"].value = name
            sheet[f"D{row}"].value = usage
            sheet[f"F{row}"].value = unit
            sheet[f"D{row2}"].value = cost
            sheet[f"D{row2 + 1}"].value = escalation

        def set_fixedcost(sheet, name, cost, escalation, index):
            row = 158 + index * 2
            if index > 5:
                return
            sheet[f"F{row}"].value = name
            sheet[f"D{row}"].value = cost
            sheet[f"D{row + 1}"].value = escalation

        def set_incentive(sheet, name, value, type, escalation, years, index):
            row = 42 + index * 3
            if index > 1:
                return
            sheet[f"F{row}"].value = name
            sheet[f"D{row}"].value = value
            sheet[f"E{row}"].value = "Tax credit" if type else "Income"
            sheet[f"D{row + 1}"].value = escalation
            sheet[f"D{row + 2}"].value = years

        h2fast_file_original = files("ProFAST.resources").joinpath("h2-fast-2022-0915.xlsm")
        shutil.copy(h2fast_file_original, filename)
        app = xw.App(visible=False)
        wb = xw.Book(filename)
        sheet = wb.sheets["Interface"]
        refurb = wb.sheets["Overrides"]
        cell_locs = {'capacity':"D26", 
         'installation cost':{
          'value': '"D66"', 
          'depr type': '"E66"', 
          'depr period': '"F66"', 
          'depreciable': '"D172"'}, 
         'non depr assets':"D67", 
         'end of proj sale non depr assets':"D68", 
         'maintenance':{'value':"D156", 
          'escalation':"D157"}, 
         'one time cap inct':{
          'value': '"D38"', 
          'depr type': '"E38"', 
          'depr period': '"F38"', 
          'depreciable': '"D174"'}, 
         'annual operating incentive':{'value':"D39", 
          'decay':"D40", 
          'sunset years':"D41"}, 
         'incidental revenue':{'value':"D52", 
          'escalation':"D53"}, 
         'commodity':{
          'name': '"D23"', 
          'unit': '"D24"', 
          'initial price': '"D29"', 
          'escalation': '"D30"'}, 
         'analysis start year':"D31", 
         'operating life':"D32", 
         'installation months':"D33", 
         'demand rampup':"D34", 
         'long term utilization':"D35", 
         'TOPC':{
          'unit price': '"D48"', 
          'decay': '"D49"', 
          'sunset years': '"D50"', 
          'support utilization': '"D51"'}, 
         'credit card fees':"D143", 
         'sales tax':"D144", 
         'road tax':{'value':"D145", 
          'escalation':"D146"}, 
         'labor':{'value':"D147", 
          'rate':"D148",  'escalation':"D149"}, 
         'license and permit':{'value':"D150", 
          'escalation':"D151"}, 
         'rent':{'value':"D152", 
          'escalation':"D153"}, 
         'property tax and insurance':"D154", 
         'admin expense':"D155", 
         'total income tax rate':"D170", 
         'capital gains tax rate':"D171", 
         'operating incentives taxable':"D173", 
         'sell undepreciated cap':"D175", 
         'tax losses monetized':"D176", 
         'tax loss carry forward years':"F177", 
         'general inflation rate':"D178", 
         'leverage after tax nominal discount rate':"D181", 
         'debt equity ratio of initial financing':"D182", 
         'debt type':"D183", 
         'loan period if used':"D184", 
         'debt interest rate':"D185", 
         'cash onhand':"D186", 
         'fraction of capital spent in constr year':"D69"}
        for i in cell_locs:
            print(i)
            if isinstance(cell_locs[i], dict):
                pf_dict = cell_locs[i].copy()
                for j in cell_locs[i]:
                    if j == "initial price":
                        pass
                    else:
                        val = self.vals[i][j]
                        if j == "depreciable":
                            val = "Yes" if val else "No"
                        sheet[cell_locs[i][j]].value = val
                        pf_dict[j] = val if j != "depreciable" else val == "Yes"
                else:
                    print(f"\t{pf_dict}")

            else:
                if i == "operating incentives taxable":
                    val = self.vals["annual operating incentive"]["taxable"]
                else:
                    if i == "fraction of capital spent in constr year":
                        val = 1
                    else:
                        val = self.vals[i]
                if i in ('tax losses monetized', 'operating incentives taxable', 'sell undepreciated cap'):
                    val = "Yes" if val else "No"
                if i == "debt type":
                    val = "Bond debt" if val == "Revolving debt" else "One time loan"
                sheet[cell_locs[i]].value = val
        else:
            i = 0
            for index, val in self.capital_items.items():
                set_capital(sheet, refurb, val.name, val.cost, val.depr_type, val.depr_period, val.refurb, i)
                i += 1
            else:
                i = 0
                for index, val in self.feedstocks.items():
                    set_feedstock(sheet, val.name, val.usage, val.cost, val.escalation, val.unit, i)
                    i += 1
                else:
                    i = 0
                    for index, val in self.coproducts.items():
                        set_coprod(sheet, val.name, val.usage, val.cost, val.escalation, val.unit, i)
                        i += 1
                    else:
                        i = 0
                        for index, val in self.incentives.items():
                            set_incentive(sheet, val.name, val.value, val.tax_credit, val.decay, val.sunset_years, i)
                            i += 1
                        else:
                            i = 0
                            for index, val in self.fixed_costs.items():
                                set_fixedcost(sheet, val.name, val.cost, val.escalation, i)
                                i += 1
                            else:
                                wb.save(filename)
                                app.quit()
                                wb.app.kill()


def depreciate(type, period, percent, cap, installation_months, MACRS_table, fyo):
    ypis = np.ceil((installation_months + 1) / 12)
    qpis = np.ceil(((installation_months + 0.5) / 12 - np.floor((installation_months + 0.5) / 12)) * 4)
    if type == "MACRS":
        col = f"Q{int(qpis)}_{period}"
        depr_table = MACRS_table[col].values
        equip_depr_sch_pct = np.concatenate((np.zeros(int(ypis) - 1), depr_table))
    else:
        if type == "Straight line":
            equip_depr_sch_pct = np.diff(np.minimum(np.cumsum(fyo / period), 1))
    equip_depr_schedule = equip_depr_sch_pct * cap
    A = percent
    B = equip_depr_schedule[equip_depr_schedule != 0]
    lenA = len(A)
    lenB = len(B)
    C, D = np.meshgrid(A, np.flip(B))
    E = (C * D).transpose()
    maxlen = max(lenA, lenB)
    rows, cols = E.shape
    rows_arr = np.arange(rows)
    cols_arr = np.arange(cols)
    diag_idx = rows_arr[(None[:None], None)] - (cols_arr - (cols - 1))
    depr_charge = np.bincount((diag_idx.ravel()), weights=(E.ravel()))
    depr_sum = sum(depr_charge)
    depr_charge = np.pad(depr_charge, (0, maxlen * 2 - len(depr_charge)))
    return (depr_charge, equip_depr_schedule)