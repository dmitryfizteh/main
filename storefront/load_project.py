#!/usr/bin/env python
# -*- coding: utf-8 -*-

#TODO: проверка, что числовые поля не могут быть пустыми
#TODO: в файле загрузке поправить все цифры привести к "руб"

import xlrd
import postgresql

# Загружаем параметры подключения к БД
f = open('passwd.txt')
for connect_str in f:
    db = postgresql.open(connect_str)
f.close()

# UNDO-скрипт
undo =''

dic = {"":None}
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

# Заполнение информации о финансовой нагрузке инициатора
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
      "fininfo_working_capital) " \
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


# создаем организацию
# TODO: проверить что такой организации еще нет в БД (и сделать это до внесения ее финданных)
id_sql = db.query("SELECT MAX(id) FROM public.organizations;");
if (id_sql[0][0] is None):
    id_org = 1
else:
    id_org = int (id_sql[0][0]) + 1

# TODO: уточнить про owner_grp (пока 1 - тестовые организации-создатели проектов)
str = "INSERT INTO public.organizations(" \
      "id, name, address, fullname, inn, kpp, ogrn, owner_grp_id, loaninfo_ebitda," \
      "loaninfo_guarantee_to_3rd, loaninfo_interest_pmts, loaninfo_interest_pmts_prj," \
      "loaninfo_total_credits, loaninfo_total_credits_with_prj, loaninfo_orders_to_3rd," \
      "fininfo_id) " \
      "VALUES ({}, '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}');".format(
    id_org, dic['Инициатор проекта'], dic['Почтовый адрес инициатора'], dic['Полное название инициатора'],dic['ИНН'],dic['КПП'],dic['ОГРН'], 1, dic['EBITDA'],
    dic['Объем выданных поручительств Заемщиком в пользу третьих лиц'],dic['Процентные платежи'],dic['Процентные платежи с учётом проекта'],
    dic['Общий долг'],dic['Общий долг с учётом проекта'],dic['Объём выданных поручений третьим лицам'], id_org_fin)
print(str)
undo = "DELETE FROM public.organizations WHERE id={0};\n".format(id_org) + undo

# TODO: создаем проект

id_sql = db.query("SELECT MAX(id) FROM public.projects;");
if (id_sql[0][0] is None):
    id_proj= 1
else:
    id_proj = int (id_sql[0][0]) + 1

#TODO: внести реальные id
development_permission_id = 1
land_right_type_id = 1
project_examination_id = 1
tech_condition_id = 1
industry_id = 1
project_type_id = 1
region_id = 77
government_support_id = 1

str = "INSERT INTO public.projects(" \
      "id, name, prerequisite, requiredlicenses, requiredlicensesdescription, " \
      "substance, technologies, dscr, financeplanperiodmonths, irrmonths, " \
      "npv, npvdiscount, payoffperiodmonths, payoffperiodmonthssimple, " \
      "firsttranche, investmentphaseperiodmonths, loanperiodmonths, " \
      "loanrate, projectpowerperiodmonths, competitivefactors, competitiveness, " \
      "organization_id, projectstatus, stage_commodities_study, stage_dev_building_info, " \
      "stage_dev_licenses, stage_equipment_supply, stage_investments_volume_own, " \
      "stage_sale_study, stage_investments_volume_total, cur_market_share_fed, " \
      "cur_market_share_reg, cur_market_volume_fed, cur_market_volume_reg, " \
      "plan_market_share_fed, plan_market_share_reg, production_volume, " \
      "development_permission_id, land_right_type_id, project_examination_id, " \
      "tech_condition_id, industry_id, project_type_id, region_id, project_total_cost, " \
      "government_support_id) " \
      "VALUES ({}, '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', " \
      "'{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}');".format(
            id_proj, dic['Наименование Проекта'], dic['Предпосылки реализации Проекта'], dic['Необходимость получения лицензий в рамках Проекта'],
            dic['Необходимый набор лицензий'], dic['Суть Проекта'], dic['Используемые в Проекте технологии'],
            dic['Средний коэффициент покрытия долга по проекту (DSCR)'], int(dic['Горизонт планирования, месяцев']), dic['IRR, %'],
            dic['NPV'], dic['Ставка дисконтирования для расчета NPV, %'], int(dic['Срок окупаемости полных инвестиционных затрат (дисконтированный)']),
            int(dic['Срок окупаемости полных инвестиционных затрат (простой)']),
            dic['Планируемая дата получения первого транша (дата)'], int(dic['Продолжительность инвестиционной фазы (мес)']), int(dic['Срок кредита (мес)']),
            dic['Ожидаемая процентная ставка, %'], int(dic['Срок выхода на проектную мощность']), dic['Факторы, обеспечивающие достижение планируемых показателей Проекта'],
            dic['Конкурентные преимущества создаваемого производства и продукции'],
            id_org, dic['Текущий статус проекта'], dic['Проработанность вопросов снабжения Проекта сырьем'], dic['Наличие строительной сметы и договора строительного подряда'],
            dic['Наличие лицензий'], dic['Наличие оборудования'], dic['собственные средства инициатора Проекта'],
            dic['Проработанность вопросов сбыта продукции Проекта'], dic['Объем вложенных средств в Проект'], dic['Текущая доля инициаторов на федеральном рынке Проекта'],
            dic['Текущая доля инициаторов на региональном рынке Проекта'], dic['Текущий объем федерального рынка'], dic['Текущий объем регионального рынка'],
            dic['Планируемая доля инициаторов на федеральном рынке'], dic['Планируемая доля инициаторов на региональном рынке'], dic['Объём производства в рамках создаваемого производства'],
            development_permission_id, land_right_type_id, project_examination_id,
            tech_condition_id, industry_id, project_type_id, region_id, dic['Общая стоимость Проекта'],
            government_support_id)
print(str)
undo = "DELETE FROM public.projects WHERE id={0};\n".format(id_proj) + undo

# TODO: связываем проект и организацию через project_initiators


# TODO: добавить документы по проекту
# TODO: создаем источники финансирования
# TODO: создаем источники поддержки
# TODO: создаем залоги
# TODO: создаем бюджет проекта
# TODO: создаем проекты в тех же отраслях
# TODO: создаем проекты в смежных отраслях