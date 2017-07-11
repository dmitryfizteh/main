#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import postgresql

# Загружаем параметры подключения к БД
f = open('passwd.txt')
for connect_str in f:
    db = postgresql.open(connect_str)
f.close()

# UNDO-скрипт
undo =''

dic = {"id":1}
rb = xlrd.open_workbook('project1.xlsx')
sheet = rb.sheet_by_index(0)
for rownum in range(2, sheet.nrows):
    row = sheet.row_values(rownum)
    if (row[0] != 0):
        dic[row[0]] = row[2]

#for key in sorted(dic.keys()):
#for key in dic.keys():
#    print("%s: %s" % (key, dic[key]))
#print (dic)

id = dic['id'] # id проекта
name = dic['Наименование Проекта']
prerequisite = dic['Предпосылки реализации Проекта']
substance = dic['Суть Проекта']
technologies = dic['Используемые в Проекте технологии']
irrmonths = dic['IRR, %'] *100
npv = dic['NPV']
npvdiscount = dic['Ставка дисконтирования для расчета NPV, %']
payoffperiodmonths = dic['Срок окупаемости полных инвестиционных затрат (дисконтированный)']
payoffperiodmonthssimple = dic['Срок окупаемости полных инвестиционных затрат (простой)']
firsttranche = dic['Планируемая дата получения первого транша (дата)']
projectstatus = dic['Текущий статус проекта']
investmentphaseperiodmonths = dic['Продолжительность инвестиционной фазы (мес)']
loanperiodmonths = dic['Срок кредита (мес)']
projectpowerperiodmonths = dic['Срок выхода на проектную мощность']
cur_market_share_fed = dic['Текущий объем федерального рынка в денежном выражении']
cur_market_share_reg = dic['Текущий объем регионального рынка в денежном выражении']
cur_market_volume_fed = dic['Текущий объем федерального рынка в натуральном выражении']
cur_market_volume_reg = dic['Текущий объем регионального рынка в натуральном выражении']
plan_market_share_fed = dic['Планируемая доля инициаторов на федеральном рынке']
plan_market_share_reg = dic['Планируемая доля инициаторов на региональном рынке']
production_volume = dic['Объём производства в рамках создаваемого производства']
dscr = dic['Средний коэффициент покрытия долга по проекту (DSCR)']
loanrate = dic['Ожидаемая процентная ставка']
requiredlicenses = dic['Необходимость получения лицензий в рамках Проекта']
competitivefactors = dic['Факторы, обеспечивающие достижение планируемых показателей Проекта']
competitiveness = dic['Конкурентные преимущества создаваемого производства и продукции']

dic[''] = None

requiredlicensesdescription = dic['']

financeplanperiodmonths = dic['']

stage_commodities_study = dic['']
stage_dev_building_info = dic['Наличие строительной сметы и договора строительного подряда']
stage_dev_licenses = dic['']
stage_equipment_supply = dic['']
stage_investments_volume_own = dic['']
stage_sale_study = dic['']
stage_investments_volume_total = dic['']

organization_id = dic['']
development_permission_id = dic['']
land_right_type_id = dic['']
project_examination_id = dic['']
tech_condition_id = dic['']
industry_id = dic['']
project_type_id = dic['']
region_id = dic['']

str = "name = {}\n prerequisite = {}\n requiredlicenses = {}\n requiredlicensesdescription = {}\n substance = {}\n technologies = {}\n dscr = {}\n financeplanperiodmonths = {}\n irrmonths = {}\n" \
      "npv = {}\n npvdiscount = {}\n payoffperiodmonths = {}\n payoffperiodmonthssimple = {}\n firsttranche = {}\n investmentphaseperiodmonths = {}\n loanperiodmonths = {}\n " \
      "loanrate = {}\n projectpowerperiodmonths = {}\n competitivefactors = {}\n competitiveness = {}\n organization_id = {}\n projectstatus = {}\n stage_commodities_study = {}\n stage_dev_building_info = {}\n " \
      "stage_dev_licenses = {}\n stage_equipment_supply = {}\n stage_investments_volume_own = {}\n stage_sale_study = {}\n stage_investments_volume_total = {}\n cur_market_share_fed = {}\n" \
      "cur_market_share_reg = {}\n cur_market_volume_fed = {}\n cur_market_volume_reg = {}\n plan_market_share_fed = {}\n plan_market_share_reg = {}\n production_volume = {}\n " \
      "development_permission_id = {}\n land_right_type_id = {}\n project_examination_id = {}\n tech_condition_id = {}\n industry_id = {}\n project_type_id = {}\n region_id = {}\n) " \
      "".format(
    name, prerequisite, requiredlicenses, requiredlicensesdescription, substance, technologies, dscr, financeplanperiodmonths, irrmonths,
    npv, npvdiscount, payoffperiodmonths, payoffperiodmonthssimple, firsttranche, investmentphaseperiodmonths, loanperiodmonths,
    loanrate, projectpowerperiodmonths, competitivefactors, competitiveness, organization_id, projectstatus, stage_commodities_study, stage_dev_building_info,
    stage_dev_licenses, stage_equipment_supply, stage_investments_volume_own, stage_sale_study, stage_investments_volume_total, cur_market_share_fed,
    cur_market_share_reg, cur_market_volume_fed, cur_market_volume_reg, plan_market_share_fed, plan_market_share_reg, production_volume,
    development_permission_id, land_right_type_id, project_examination_id, tech_condition_id, industry_id, project_type_id, region_id)
#print(str)

# TODO: создаем проект
str = "INSERT INTO public.projects(id, name, prerequisite, requiredlicenses, requiredlicensesdescription, substance, technologies, dscr, financeplanperiodmonths, irrmonths," \
      "npv, npvdiscount, payoffperiodmonths, payoffperiodmonthssimple, firsttranche, investmentphaseperiodmonths, loanperiodmonths, " \
      "loanrate, projectpowerperiodmonths, competitivefactors, competitiveness, organization_id, projectstatus, stage_commodities_study, stage_dev_building_info, " \
      "stage_dev_licenses, stage_equipment_supply, stage_investments_volume_own, stage_sale_study, stage_investments_volume_total, cur_market_share_fed," \
      "cur_market_share_reg, cur_market_volume_fed, cur_market_volume_reg, plan_market_share_fed, plan_market_share_reg, production_volume, " \
      "development_permission_id, land_right_type_id, project_examination_id, tech_condition_id, industry_id, project_type_id, region_id) " \
      "VALUES ({}, '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}'" \
      ", '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}');".format(
    id, name, prerequisite, requiredlicenses, requiredlicensesdescription, substance, technologies, dscr, financeplanperiodmonths, irrmonths,
    npv, npvdiscount, payoffperiodmonths, payoffperiodmonthssimple, firsttranche, investmentphaseperiodmonths, loanperiodmonths,
    loanrate, projectpowerperiodmonths, competitivefactors, competitiveness, organization_id, projectstatus, stage_commodities_study, stage_dev_building_info,
    stage_dev_licenses, stage_equipment_supply, stage_investments_volume_own, stage_sale_study, stage_investments_volume_total, cur_market_share_fed,
    cur_market_share_reg, cur_market_volume_fed, cur_market_volume_reg, plan_market_share_fed, plan_market_share_reg, production_volume,
    development_permission_id, land_right_type_id, project_examination_id, tech_condition_id, industry_id, project_type_id, region_id)

print(str)

#TODO: заменить id на запрос max id
id_sql = db.query("SELECT MAX(id) FROM public.organization_finances;");
if (id_sql[0][0] is None):
    id_org_fin = 1
else:
    id_org_fin = int (id_sql[0][0]) + 1

str = "INSERT INTO public.organization_finances(" \
      "id, fininfo_accounts_payable, fininfo_balance_currency, fininfo_profit_before_tax," \
      "fininfo_current_assets, fininfo_deprecation, fininfo_deprecation_np, " \
      "fininfo_deprecation_np_4q, fininfo_exch_rate_diffs, fininfo_inprog_development, " \
      "fininfo_lease_obligations, fininfo_leasing_payments, fininfo_credits_long, " \
      "fininfo_loans_long, fininfo_longterm_debt, fininfo_investments_longterm, " \
      "fininfo_main_assets, fininfo_net_profit, fininfo_onetime_incomes, " \
      "fininfo_own_funds, fininfo_prof_ebitda_pct, fininfo_prof_ebitda_pct_4q, " \
      "fininfo_prof_np, fininfo_requested_credit, fininfo_revenue, fininfo_revenue_4q, " \
      "fininfo_sales_profit, fininfo_credits_short, fininfo_loans_short, " \
      "fininfo_investments_shortterm, fininfo_subsidies, fininfo_tot_liquidity_ratio, " \
      "fininfo_working_capital)" \
      "VALUES ({}, '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', " \
      "'{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}','{}', '{}', '{}','{}', '{}', '{}', '{}');".format(
    id_org_fin, dic['Кредиторская задолженность'], dic['Валюта Баланса'], dic['Прибыль до налогообложения'],
    dic['Оборотные активы'], dic['Амортизация (А)'], dic['ЧП + А за период'],
    dic['ЧП + А (за последние 4 квартала)'], dic['Курсовые разницы (полож. минус отриц.), за период'], dic['Незавершенное производство'],
    dic['Обязательства по лизингу'], dic['Лизинговые платежи в составе себестоимости, за период'], dic['Долгосрочные кредиты'],
    dic['Долгосрочные займы'], dic['Долгосрочный долг/(ЧП+А)'], dic['Долгосрочные финансовые вложения'],
    dic['Основные средства (тыс руб)'], dic['Чистая прибыль (ЧП)'], dic['Единоразовые доходы/расходы, за исключением субсидий, (доходы минус расходы) - за период'],
    dic['Собственные средства/Валюта баланса, %'], dic['Рентабельность EBITDA, % - за период'], dic['Рентабельность EBITDA, % - за последние 4 кв.'],
    dic['Рентабельность ЧП, % - за период'], dic['Испрашиваемый кредит'], dic['Выручка'], dic['Выручка (за последние 4 квартала)'],
    dic['Прибыль от продаж'], dic['Краткосрочные кредиты'], dic['Краткосрочные займы'],
    dic['Краткосрочные финансовые вложения'], dic['Субсидии полученные, за период'], dic['Коэффициент общей ликвидности'],
    dic['Собственные оборотные средства'])
print(str)

undo = "DELETE FROM public.organization_finances WHERE id={0};\n".format(id_org_fin) + undo

# TODO: создаем организацию
# TODO: связываем проект и организацию через project_initiators
# TODO: создаем источники финансирования
# TODO: создаем источники поддержки
# TODO: создаем залоги
# TODO: создаем стоимость проекта
# TODO: создаем проекты в тех же отраслях
# TODO: создаем проекты в смежных отраслях
