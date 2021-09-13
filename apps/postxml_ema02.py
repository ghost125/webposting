# -*-coding: utf-8-*-
# @Project:posting
# @File   :postxml_ema01.py
# @Date   :2021-03-05
# @Author  :GDCPL

# -*-coding: utf-8-*-
# @Project:posting
# @File   :postxml_ema.py
# @Date   :2021-02-04
# @Author  :GDCPL

import hashlib

import re

from lxml import etree
import pandas as pd
import warnings
from collections import OrderedDict

warnings.filterwarnings('ignore')


def prettyXml(element, indent, newline, level=0):
    if element:
        if element.text is None or not element.text.isspace():
            element.text = newline + indent * (level + 1)
        else:
            element.text = newline + indent * (level + 1) + element.text + newline + indent * (level + 1)

    temp = list(element)
    for subelement in temp:
        if temp.index(subelement) < (len(temp) - 1):
            subelement.tail = newline + indent * (level + 1)
        else:
            subelement.tail = newline + indent * level
        prettyXml(subelement, indent, newline, level=level + 1)


def group_id(instr):
    m = hashlib.md5()
    m.update(str(instr).encode())
    return m.hexdigest() + '+' + '_'.join(str(instr).split(' '))


def ret_not_null(value, first=True):
    if pd.isna(value):
        return None
    elif value is None:
        return None
    else:
        if isinstance(value, int):
            vv = str(value)
        elif isinstance(value, float):
            if value == int(value):
                vv = str(int(value))
            else:
                vv = str(value)
        elif isinstance(value, str):
            try:
                vv = int(value)
            except ValueError:
                try:
                    vv = float(value)
                    if vv == int(vv):
                        vv = str(int(vv))
                    else:
                        vv = str(vv)
                except ValueError:
                    vv = value.strip()
                    if re.search('(\d+\.?\d*)\s?\(?.*N\s*=\s*(\d+)?', vv) is not None:
                        vv = re.search('(\d+\.?\d*)\s?\(?.*N\s*=\s*(\d+)?', vv).groups()
                        if first and isinstance(vv, tuple):
                            vv = vv[0]
                    elif re.search('(\d+\.?\d*)\s?\(', vv) is not None:
                        vv = re.search('(\d+\.?\d*)\s?\(', vv).groups()
                        if first and isinstance(vv, tuple):
                            vv = vv[0]
        else:
            vv = value
        return vv



def df_rename(indf):
    tmp = {i: str(i).lower() for i in indf.columns}
    indf.reset_index(drop=True, inplace=True)
    indf.rename(tmp, axis=1, inplace=True)


def prettyxmlfile(infile):
    root = etree.parse(infile).getroot()
    prettyXml(root, '\t', '\n')
    obj_xml = etree.tostring(root, pretty_print=True, xml_declaration=True)
    with open(infile, "w") as xml_writer:
        xml_writer.write(bytes.decode(obj_xml, 'UTF-8'))


# prettyxmlfile(r'C:\Studies\posting\pilot\19773\19773_PharmaCM ERF_Original_v.0.16_2021-04-26.xml')

class postxml_ema(object):
    def __init__(self, inxml):
        self.root = etree.parse(inxml).getroot()
        self.xsi = "http://www.w3.org/2001/XMLSchema-instance"

    def get_discon_reason(self):
        reason_ids=self.root.xpath('subjectDisposition/reasonsNotCompleted/reasonNotCompleted')
        all_reasons=[]
        for reason in reason_ids:
            _reas=[]
            _reas.append(reason.attrib['id'])
            if reason.xpath('type/value'):
                _reas.append(reason.xpath('type/value')[0].text)
            else:
                _reas.append('')
            if reason.xpath('otherReason'):
                _reas.append(reason.xpath('otherReason')[0].text)
            else:
                _reas.append('')
            all_reasons.append(_reas[:])
        return all_reasons

    def get_category(self,element='//baselineCharacteristics/studyCategoricalCharacteristic'):
        cats=self.root.xpath(element)
        all_cats={}
        for cat in cats:
            all_cat = []
            _p_title=cat.xpath('title')[0].text
            _cats=cat.xpath('categories')[0]
            for _cat in _cats.xpath('category'):
                _all_cat=[]
                _all_cat.append(_cat.attrib['id'])
                _all_cat.append(_cat.xpath('name')[0].text)
                all_cat.append(_all_cat[:])
            all_cats[_p_title]=all_cat[:]
        return all_cats



    def get_baseline_groups(self):
        all_grps = []
        cols = list(self.root.xpath('//baselineReportingGroups/baselineReportingGroup')[0].attrib.keys())
        for grp in self.root.xpath('//baselineReportingGroups/baselineReportingGroup'):
            _temp = []
            for i in grp.attrib:
                _temp.append(grp.attrib[i])
            if grp.xpath('description'):
                if 'description' not in cols:
                    cols.append('description')
                _temp.append(grp.xpath('description')[0].text)

            all_grps.append(_temp[:])
        return pd.DataFrame(all_grps, columns=cols)

    def get_particpate_groups(self):
        all_grps = []
        cols = list(self.root.xpath('//postAssignmentPeriod/arms/arm')[0].attrib.keys())
        for grp in self.root.xpath('//postAssignmentPeriod/arms/arm'):
            _temp = []
            for i in grp.attrib:
                _temp.append(grp.attrib[i])
            if grp.xpath('description'):
                if 'description' not in cols:
                    cols.append('description')
                _temp.append(grp.xpath('description')[0].text)
            if grp.xpath('title'):
                if 'title' not in cols:
                    cols.append('title')
                _temp.append(grp.xpath('title')[0].text)

            all_grps.append(_temp[:])
        return pd.DataFrame(all_grps, columns=cols)

    def get_outcome_groups(self):
        all_grps = []
        cols = list(self.root.xpath('//armReportingGroups/armReportingGroup')[0].attrib.keys())
        for grp in self.root.xpath('//armReportingGroups/armReportingGroup'):
            _temp = []
            for i in grp.attrib:
                _temp.append(grp.attrib[i])
            if grp.xpath('description'):
                if 'description' not in cols:
                    cols.append('description')
                _temp.append(grp.xpath('description')[0].text)
            if grp.xpath('title'):
                if 'title' not in cols:
                    cols.append('title')
                _temp.append(grp.xpath('title')[0].text)

            all_grps.append(_temp[:])
        return pd.DataFrame(all_grps, columns=cols)

    def get_analysissets(self):
        return [i.attrib['id'] for i in self.root.xpath('subjectAnalysisSets/subjectAnalysisSet')]

    def addstudyno(self, studyno):
        study = self.root.xpath("//sponsorProtocolCode")
        if study:
            study[0].text = studyno

    def add_agegroup(self, agegroupinput, column):
        '''
        add age group information to
        :param agegroupinput: input dataframe, agegroup,count 2 variables must be exist
        :return:
        '''
        df_rename(agegroupinput)
        column = column.lower()
        agegroup = self.root.xpath("//trialInformation/populationAgeGroup")
        if agegroup:
            for element in agegroup[0].getchildren():
                agegroup[0].remove(element)
            inagegroups = agegroupinput[['agegroup', column]].to_dict(orient='records')
            ets = OrderedDict(
                {1: 'inUtero', 2: 'pretermNewbornInfants', 3: 'newborns', 4: 'infantsAndToddlers', 5: 'children',
                 6: 'adolescents', 7: 'adults', 8: 'elderly65To84', 9: 'elderlyOver85'})
            inagedict = {int(ret_not_null(i['agegroup'])): ret_not_null(i[column]) for i in inagegroups}
            for i in ets:
                et = etree.Element(ets[i])
                txt = inagedict.get(i)
                if not pd.isna(txt):
                    et.text = str(ret_not_null(txt))
                else:
                    et.text = '0'
                agegroup[0].append(et)

    def add_patientflow(self, indf, groups):
        df_rename(indf)
        groups = [i.lower() for i in groups]
        post_periods = self.root.xpath("//postAssignmentPeriod")
        for np, period in enumerate(post_periods):
            used_df = indf[(indf['period'] == np + 1)|(indf['period'] == 999)]
            if period.xpath('arms/arm'):
                for ai, arm in enumerate(period.xpath('arms/arm')):
                    used_arm={i['categoryid']:i[groups[ai]] for i in used_df.to_dict(orient='records') if i['period'] !=999}
                    milestones = ['startedMilestoneAchievement', 'completedMilestoneAchievement',
                                  'otherMilestoneAchievements/otherMilestoneAchievement']
                    for mstone in milestones:
                        if arm.xpath(mstone):
                            for c_et in arm.xpath(mstone):
                                c_et_dict=c_et.attrib
                                for key in c_et_dict.values():
                                    if key in used_arm.keys():
                                        try:
                                            c_et.xpath('subjects')[0].text = str(
                                            ret_not_null(used_arm[key]))
                                        except:
                                            etree.SubElement(c_et,'subjects').text=str(
                                            ret_not_null(used_arm[key]))
                    used_arm = {i['categoryid']: i[groups[ai]] for i in used_df.to_dict(orient='records') if
                                i['period'] == 999}
                    notcompreasons=arm.xpath('notCompletedReasonDetails')[0]
                    for n_et in notcompreasons.getchildren():
                        notcompreasons.remove(n_et)
                    for reasonid in used_arm.keys():
                        if int(ret_not_null(used_arm[reasonid]))>0:
                            reas_et=etree.SubElement(notcompreasons,'reasonDetail',{'reasonNotCompletedId':reasonid})
                            etree.SubElement(reas_et,'subjects').text=str(ret_not_null(used_arm[reasonid]))

        #drop not needed reasons
        used_ids = set(indf['categoryid'])
        p_reason=self.root.xpath('subjectDisposition/reasonsNotCompleted')[0]
        for d_reason in self.root.xpath('subjectDisposition/reasonsNotCompleted/reasonNotCompleted'):
            if d_reason.attrib['id'] not in used_ids:
                p_reason.remove(d_reason)



    def set_baseline_groups(self, indf):
        df_rename(indf)
        indict = {indf.loc[i, 'armid']: indf.loc[i, 'n'] for i in range(indf.shape[0])}
        for grp in self.root.xpath('//baselineReportingGroups/baselineReportingGroup'):
            nsubj = grp.xpath('subjects')[0]
            if grp.attrib['armId'] in indict.keys():
                nsubj.text = str(indict[grp.attrib['armId']])

    def _set_baseline_groups(self, indf,groups):
        df_rename(indf)
        inlist = [indf.loc[0,grp.lower()] for grp in groups]
        for grp_i, grp in enumerate(self.root.xpath('//baselineReportingGroups/baselineReportingGroup')):
            nsubj = grp.xpath('subjects')[0]
            try:
                nsubj.text=str(ret_not_null(inlist[grp_i],first=False)[1])
            except:
                pass


    def set_analysissets(self, indf):
        df_rename(indf)
        temp_dict=dict(zip(indf['id'],indf['subjects']))
        ets = self.root.xpath('subjectAnalysisSets/subjectAnalysisSet')
        for et in enumerate(ets):
            if et.attrib['id'] in temp_dict:
                for et1 in et.getchildren():
                    if et1.tag == 'subjects':
                        et1.text = str(ret_not_null(temp_dict[et.attrib['id']]))

    def add_cat_character(self, element, indf, groups, total=None):
        df_rename(indf)
        groups = [i.lower() for i in groups]
        rep_grps = element.xpath('reportingGroups/reportingGroup')
        for i, grp in enumerate(rep_grps):
            indict={j['categoryid']:j[groups[i]] for j in indf[['categoryid',groups[i]]].to_dict(orient='records')}
            coutvalues = grp.xpath('countableValues/countableValue')
            for j in coutvalues:
                if j.attrib['categoryId'] in indict.keys():
                    if j.xpath('value'):
                        j.xpath('value')[0].attrib.clear()
                        j.xpath('value')[0].text=str(ret_not_null(indict[j.attrib['categoryId']]))
                    else:
                        etree.SubElement(j,'value').text=str(ret_not_null(indict[j.attrib['categoryId']]))

        if total is not None:
            total_ctvals = element.xpath('totalBaselineGroup/countableValues/countableValue')
            indict = {j['categoryid']: j[total] for j in indf[['categoryid', total]].to_dict(orient='records')}
            for cet in total_ctvals:
                if cet.attrib['categoryId'] in indict.keys():
                    if cet.xpath('value'):
                        cet.xpath('value')[0].attrib.clear()
                        cet.xpath('value')[0].text=str(ret_not_null(indict[cet.attrib['categoryId']]))
                    else:
                        etree.SubElement(cet,'value').text=str(ret_not_null(indict[cet.attrib['categoryId']]))

        #delete not needed category
        all_cats=set(indf['categoryid'])
        for cat in element.xpath('categories/category'):
            if cat.attrib['id'] not in all_cats:
                element.xpath('categories')[0].remove(cat)

    def add_cat_character2(self, element, indf, groups, total=None,asets=None):
        df_rename(indf)
        groups = [i.lower() for i in groups]
        if element.xpath('reportingGroups/reportingGroup'):
            rep_grps = element.xpath('reportingGroups/reportingGroup')
            for i, grp in enumerate(rep_grps):
                self.add_baseline_report_value(grp,indf,groups[i])
        else:
            if element.xpath('subjectAnalysisSets/subjectAnalysisSet'):
                if asets is not None:
                    rep_grps = element.xpath('reportingGroups/reportingGroup')
                    for i, grp in enumerate(rep_grps):
                        self.add_baseline_report_value(grp, indf, asets[i])

        if total is not None:
            total_ctvals = element.xpath('totalBaselineGroup')
            for cet in total_ctvals:
                self.add_baseline_report_value(cet,indf,total)

        #delete not needed category
        all_cats=set(indf['categoryid'])
        for cat in element.xpath('categories/category'):
            if cat.attrib['id'] not in all_cats:
                element.xpath('categories')[0].remove(cat)
        for set_ in element.xpath('subjectAnalysisSets/subjectAnalysisSet'):
            for cat in set_.xpath('countableValues/countableValue'):
                if cat.attrib['categoryId'] not in all_cats:
                    cat.getparent().remove(cat)

    def add_baseline_report_value(self,element,indf,groupset):
        def _inner():
            if 'categoryId' in cat.attrib:
                cc_df = indf[indf['categoryid'] == cat.attrib['categoryId']]
                if cc_df.empty:
                    cat.getparent().remove(cat)
            else:
                cc_df = indf.copy()
            temp_dict = dict(zip(cc_df['stats'], cc_df[groupset]))
            if 'n' in temp_dict.keys() and cat.tag == 'countableValue':
                if cat.xpath('value'):
                    cat.xpath('value')[0].attrib.clear()
                    cat.xpath('value')[0].text = str(ret_not_null(temp_dict['n']))
                else:
                    etree.SubElement(cat, 'value').text = str(ret_not_null(temp_dict['n']))
            if 'mean' in temp_dict and cat.tag == 'tendencyValue':
                if cat.xpath('value'):
                    cat.xpath('value')[0].attrib.clear()
                    cat.xpath('value')[0].text = str(ret_not_null(temp_dict['mean']))
                else:
                    etree.SubElement(element, 'value').text = str(ret_not_null(temp_dict['mean']))
            if 'sd' in temp_dict and cat.tag == 'dispersionValue':
                if cat.xpath('value'):
                    cat.xpath('value')[0].attrib.clear()
                    cat.xpath('value')[0].text = str(ret_not_null(temp_dict['sd']))
                else:
                    etree.SubElement(element, 'value').text = str(ret_not_null(temp_dict['sd']))
            if 'max' in temp_dict and cat.tag == 'dispersionValue':
                if cat.xpath('highRangeValue'):
                    cat.xpath('highRangeValue')[0].attrib.clear()
                    cat.xpath('highRangeValue')[0].text = str(ret_not_null(temp_dict['max']))
                else:
                    etree.SubElement(element, 'highRangeValue').text = str(ret_not_null(temp_dict['max']))

        if element.xpath('countableValues/countableValue'):
            for cat in element.xpath('countableValues/countableValue'):
                _inner()
        if element.xpath('tendencyValue'):
            for cat in element.xpath('tendencyValue'):
                _inner()

        if element.xpath('dispersionValue'):
            for cat in element.xpath('dispersionValue'):
                _inner()

    def add_baseline_cat_characters(self, indf, groups,total=None,asets=None):
        df_rename(indf)
        groups = [i.lower() for i in groups]
        baseline_cat = self.root.xpath('//studyCategoricalCharacteristic')
        for bic, bcat in enumerate(baseline_cat):
            _df = indf[indf['_nr_'] == bic + 1].reset_index(drop=True)
            #self.add_cat_character(bcat, _df, groups,total)
            self.add_cat_character2(bcat,_df,groups,total,asets)
        if asets is None:
            self._set_baseline_groups(indf,groups)


    def add_age_cat_characters(self, indf, groups,total=None,asets=None):
        df_rename(indf)
        groups = [i.lower() for i in groups]
        baseline_cat = self.root.xpath('//ageCategoricalCharacteristic')
        for bic, bcat in enumerate(baseline_cat):
            _df = indf[indf['_nr_'] == bic + 1]
            #self.add_cat_character(bcat, _df, groups,total)
            self.add_cat_character2(bcat, _df, groups, total,asets)
        #self._set_baseline_groups(indf, groups)


    def add_gender_cat_characters(self, indf, groups, total=None,asets=None):
        df_rename(indf)
        groups = [i.lower() for i in groups]
        baseline_cat = self.root.xpath('//genderCategoricalCharacteristic')
        for bic, bcat in enumerate(baseline_cat):
            _df = indf[indf['_nr_'] == bic + 1]
            #self.add_cat_character(bcat, _df, groups,total)
            self.add_cat_character2(bcat, _df, groups, total,asets)
        if asets is None:
            self._set_baseline_groups(indf, groups)

    def add_cont_character(self, element, indf, groups):
        df_rename(indf)
        groups = [i.lower() for i in groups]
        rep_grps = element.xpath('reportingGroups/reportingGroup')
        for ig, grp in enumerate(rep_grps):
            tendvalue = grp.xpath('tendencyValue')[0]
            dispersionvalue = grp.xpath('dispersionValue')[0]

            indf_dict = {indf.loc[i, 'categoryid'].lower(): indf.loc[i, groups[ig]] for i in range(indf.shape[0])}
            if 'mean' in indf_dict.keys():
                if tendvalue.xpath('value'):
                    tendvalue.xpath('value')[0].attrib.clear()
                    tendvalue.xpath('value')[0].text=ret_not_null(indf_dict['mean'])
                else:
                    etree.SubElement(tendvalue, 'value').text = ret_not_null(indf_dict['mean'])
            if 'sd' in indf_dict.keys():
                if dispersionvalue.xpath('value'):
                    dispersionvalue.xpath('value')[0].attrib.clear()
                    dispersionvalue.xpath('value')[0].text = ret_not_null(indf_dict['sd'])
                else:
                    etree.SubElement(dispersionvalue, 'value').text = ret_not_null(indf_dict['sd'])

            if 'max' in indf_dict.keys():
                if dispersionvalue.xpath('highRangeValue'):
                    dispersionvalue.xpath('highRangeValue')[0].attrib.clear()
                    dispersionvalue.xpath('highRangeValue')[0].text = ret_not_null(indf_dict['max'])
                else:
                    etree.SubElement(dispersionvalue, 'highRangeValue').text = ret_not_null(indf_dict['max'])

    def add_cont_character2(self, element, indf, groups,asets=None):
        df_rename(indf)
        groups = [i.lower() for i in groups]
        if element.xpath('reportingGroups/reportingGroup'):
            rep_grps = element.xpath('reportingGroups/reportingGroup')
            for ig, grp in enumerate(rep_grps):
                self.add_baseline_report_value(grp,indf,groups[ig])
        else:
            if element.xpath('subjectAnalysisSets/subjectAnalysisSet'):
                if asets is not None:
                    rep_grps = element.xpath('reportingGroups/reportingGroup')
                    for i, grp in enumerate(rep_grps):
                        self.add_baseline_report_value(grp, indf, asets[i])

    def add_baseline_cont_characters(self, indf, groups,asets=None):
        df_rename(indf)
        groups = [i.lower() for i in groups]
        baseline_cont = self.root.xpath('//studyContinuousCharacteristic')
        for bic, bcont in enumerate(baseline_cont):
            _df = indf[indf['_nr_'] == bic + 1]
            #self.add_cont_character(bcont, _df, groups)
            self.add_cont_character2(bcont,_df,groups,asets)

        #self._set_baseline_groups(indf,groups)

    def add_age_cont_characters(self, indf, groups,asets=None):
        df_rename(indf)
        groups = [i.lower() for i in groups]
        age_cont = self.root.xpath('//ageContinuousCharacteristic')
        for bic, bcont in enumerate(age_cont):
            _df = indf[indf['_nr_'] == bic + 1]
            #self.add_cont_character(bcont, _df, groups)
            self.add_cont_character2(bcont, _df, groups,asets)

        #self._set_baseline_groups(indf, groups)

    def add_report_value(self, element, type_, indf, group):
        df_rename(indf)

        temp_dict = {indf.loc[i, 'categoryid']: indf.loc[i, group] for i in range(indf.shape[0])}
        if type_.upper() == 'TRUE':
            for cat in element.xpath('countableValues/countableValue'):
                if cat.attrib['categoryId'] in temp_dict.keys():
                    if cat.xpath('value'):
                        cat.xpath('value')[0].text=str(ret_not_null(temp_dict[cat.attrib['categoryId']]))
                    else:
                        etree.SubElement(cat,'value').text=str(ret_not_null(temp_dict[cat.attrib['categoryId']]))

        else:
            et1 = element.xpath('tendencyValues/tendencyValue')[0]
            if 'mean' in temp_dict.keys():
                if not pd.isna(temp_dict['mean']):
                    if et1.xpath('value'):
                        et1.xpath('value')[0].attrib.clear()
                        et1.xpath('value')[0].text=str(ret_not_null(temp_dict['mean']))
                    else:
                        etree.SubElement(et1,'value').text= str(ret_not_null(temp_dict['mean']))

            et2 = element.xpath('dispersionValues/dispersionValue')[0]
            if 'sd' in temp_dict.keys():
                if not pd.isna(temp_dict['sd']):
                    if et2.xpath('value'):
                        et2.xpath('value')[0].attrib.clear()
                        et2.xpath('value')[0].text= str(ret_not_null(temp_dict['sd']))
                    else:
                        etree.SubElement(et2,'value').text=str(ret_not_null(temp_dict['sd']))
            if 'max' in temp_dict.keys():
                if not pd.isna(temp_dict['max']):
                    if et2.xpath('highRangeValue'):
                        et2.xpath('highRangeValue')[0].attrib.clear()
                        et2.xpath('highRangeValue')[0].text= str(ret_not_null(temp_dict['max']))
                    else:
                        etree.SubElement(et2,'highRangeValue').text=str(ret_not_null(temp_dict['max']))

        if str(indf.loc[0, group]).find('N') >= 0:
            N = ret_not_null(indf.loc[0, group], first=False)[1]
            if element.xpath('subjects'):
                et = element.xpath('subjects')[0]
                et.text = str(ret_not_null(N))
            else:
                etree.SubElement(element,'subjects').text=str(ret_not_null(N))

    def add_report_value2(self, element, indf, groupset):
        df_rename(indf)
        def _inner():
            if 'categoryId' in cat.attrib:
                cc_df = indf[indf['categoryid'] == cat.attrib['categoryId']]
                if cc_df.empty:
                    cat.getparent().remove(cat)
            else:
                cc_df = indf.copy()
            temp_dict = dict(zip(cc_df['stats'], cc_df[groupset]))
            if 'n' in temp_dict.keys() and cat.tag == 'countableValue':
                if cat.xpath('value'):
                    cat.xpath('value')[0].attrib.clear()
                    cat.xpath('value')[0].text = str(ret_not_null(temp_dict['n']))
                else:
                    etree.SubElement(cat, 'value').text = str(ret_not_null(temp_dict['n']))
            if 'mean' in temp_dict and cat.tag == 'tendencyValue':
                if cat.xpath('value'):
                    cat.xpath('value')[0].attrib.clear()
                    cat.xpath('value')[0].text = str(ret_not_null(temp_dict['mean']))
                else:
                    etree.SubElement(element, 'value').text = str(ret_not_null(temp_dict['mean']))
            if 'sd' in temp_dict and cat.tag == 'dispersionValue':
                if cat.xpath('value'):
                    cat.xpath('value')[0].attrib.clear()
                    cat.xpath('value')[0].text = str(ret_not_null(temp_dict['sd']))
                else:
                    etree.SubElement(element, 'value').text = str(ret_not_null(temp_dict['sd']))
            if 'max' in temp_dict and cat.tag == 'dispersionValue':
                if cat.xpath('highRangeValue'):
                    cat.xpath('highRangeValue')[0].attrib.clear()
                    cat.xpath('highRangeValue')[0].text = str(ret_not_null(temp_dict['max']))
                else:
                    etree.SubElement(element, 'highRangeValue').text = str(ret_not_null(temp_dict['max']))
        if element.xpath('countableValues/countableValue'):
            for cat in element.xpath('countableValues/countableValue'):
                _inner()
        if element.xpath('tendencyValues/tendencyValue'):
            for cat in element.xpath('tendencyValues/tendencyValue'):
                _inner()

        if element.xpath('dispersionValues/dispersionValue'):
            for cat in element.xpath('dispersionValues/dispersionValue'):
                _inner()

        if str(indf.loc[0, groupset]).find('N') >= 0:
            N = ret_not_null(indf.loc[0, groupset], first=False)[1]
            if element.xpath('subjects'):
                et = element.xpath('subjects')[0]
                et.text = str(ret_not_null(N))
            else:
                etree.SubElement(element,'subjects').text=str(ret_not_null(N))


    def add_statistical_analysis(self, element, indf):
        df_rename(indf)
        ana_ets = element.xpath('statisticalAnalyses/statisticalAnalysis')
        temp_dict = indf.to_dict(orient='records')
        for ana_i,ana_et in enumerate(ana_ets):
            if 'title' in temp_dict[ana_i].keys():
                if ana_et.xpath('title'):
                    if ana_et.xpath('title')[0].text != temp_dict[ana_i]['title']:
                        ana_et.xpath('title')[0].text = temp_dict[ana_i]['title']
                else:
                    etree.SubElement(ana_et,'title').text=temp_dict[ana_i]['title']
            if 'description' in temp_dict[ana_i].keys():
                if ana_et.xpath('description'):
                    if ana_et.xpath('description')[0].text != temp_dict[ana_i]['description']:
                        ana_et.xpath('description')[0].text = temp_dict[ana_i]['description']
                else:
                    etree.SubElement(ana_et, 'description').text = temp_dict[ana_i]['description']
            if 'othermethod' in temp_dict[ana_i].keys():
                if ana_et.xpath('statisticalHypothesisTest/otherMethod'):
                    if ana_et.xpath('statisticalHypothesisTest/otherMethod')[0].text != temp_dict[ana_i]['othermethod']:
                        ana_et.xpath('statisticalHypothesisTest/otherMethod')[0].text = temp_dict[ana_i]['othermethod']
                else:
                    if ana_et.xpath('statisticalHypothesisTest'):
                        etree.SubElement(ana_et.xpath('statisticalHypothesisTest')[0],'otherMethod').text=temp_dict[ana_i]['othermethod']
                    else:
                        sht=etree.SubElement(ana_et,'statisticalHypothesisTest')
                        etree.SubElement(sht,'otherMethod').text= temp_dict[ana_i]['othermethod']
            used_dict=temp_dict[ana_i]
            pvalue=ana_et.xpath('statisticalHypothesisTest/value')
            if 'pvalue' in used_dict.keys():
                if len(pvalue)>0:
                    pvalue[0].attrib.clear()
                    pvalue[0].text = str(used_dict['pvalue']).split('<')[1] if str(used_dict['pvalue']).find('<') >= 0 \
                        else str(used_dict['pvalue'])
                else:
                    if ana_et.xpath('statisticalHypothesisTest'):
                        etree.SubElement(ana_et.xpath('statisticalHypothesisTest')[0],'value').text= \
                            str(used_dict['pvalue']).split('<')[1] if str(used_dict['pvalue']).find('<') >= 0 \
                                else str(used_dict['pvalue'])

                    else:
                        sht=etree.SubElement(ana_et,'statisticalHypothesisTest')
                        etree.SubElement(sht,'value').text=str(used_dict['pvalue']).split('<')[1] \
                            if str(used_dict['pvalue']).find('<') >= 0 else str(used_dict['pvalue'])

                pval_opt=ana_et.xpath('statisticalHypothesisTest/valueEqualityRelation')

                if len(pval_opt)>0:
                    pval_opt[0].text = '<' if str(used_dict['pvalue']).find('<') >= 0 else '='
                else:
                    if ana_et.xpath('statisticalHypothesisTest'):
                        etree.SubElement(ana_et.xpath('statisticalHypothesisTest')[0], 'valueEqualityRelation').text = \
                            '<' if str(used_dict['pvalue']).find('<') >= 0 else '='
                    else:
                        sht = etree.SubElement(ana_et, 'statisticalHypothesisTest')
                        etree.SubElement(sht, 'valueEqualityRelation').text = '<' if str(used_dict['pvalue']).find('<') >= 0 else '='

            if 'lowerlimit' in used_dict.keys() or 'upperlimit' in used_dict.keys():
                interval=ana_et.xpath('parameterEstimate/confidenceInterval')
                if len(interval)>0:
                    interval=interval[0]
                    if pd.notna(used_dict['lowerlimit']):
                        if interval.xpath('lowerLimit'):
                            interval.xpath('lowerLimit')[0].attrib.clear()
                            interval.xpath('lowerLimit')[0].text=str(ret_not_null(used_dict['lowerlimit']))
                        else:
                            lowerL=etree.SubElement(interval,'lowerLimit')
                            lowerL.text=str(ret_not_null(used_dict['lowerlimit']))
                            interval.insert(0,lowerL)
                    if pd.notna(used_dict['percentage']):
                        if interval.xpath('percentage'):
                            interval.xpath('percentage')[0].attrib.clear()
                            interval.xpath('percentage')[0].text=str(ret_not_null(used_dict['percentage']))
                        else:
                            percent=etree.SubElement(interval,'percentage')
                            percent.text=str(ret_not_null(used_dict['percentage']))
                            interval.insert(1,percent)
                    if pd.notna(used_dict['upperlimit']):
                        if interval.xpath('upperLimit'):
                            interval.xpath('upperLimit')[0].attrib.clear()
                            interval.xpath('upperLimit')[0].text=str(ret_not_null(used_dict['upperlimit']))
                        else:
                            upperL=etree.SubElement(interval,'upperLimit')
                            upperL.text=str(ret_not_null(used_dict['upperlimit']))
                            interval.insert(-1,upperL)

            if 'effectestimate' in used_dict.keys():
                eE_list=ana_et.xpath('parameterEstimate/effectEstimate')
                if len(eE_list)>0:
                    eE = eE_list[0]
                    if pd.notna(used_dict['effectestimate']):
                        eE.attrib.clear()
                        eE.text=str(ret_not_null(used_dict['effectestimate']))

            if 'othertype' in used_dict.keys():
                eE_list=ana_et.xpath('parameterEstimate/otherType')
                if len(eE_list)>0:
                    eE = eE_list[0]
                    if pd.notna(used_dict['othertype']):
                        eE.attrib.clear()
                        eE.text=str(ret_not_null(used_dict['othertype']))

            if 'pointestimate' in used_dict.keys():
                eE_list=ana_et.xpath('parameterEstimate/pointEstimate')
                if len(eE_list)>0:
                    eE = eE_list[0]
                    if pd.notna(used_dict['pointestimate']):
                        eE.attrib.clear()
                        eE.text=str(ret_not_null(used_dict['pointestimate']))

            if 'variabilityestimate' in used_dict.keys():
                eE_list=ana_et.xpath('parameterEstimate/variabilityEstimate')
                if len(eE_list)>0:
                    eE=eE_list[0]
                    if len(eE.xpath('dispersionValue'))>0:
                        eE_ = eE.xpath('dispersionValue')[0]
                        if pd.notna(used_dict['variabilityestimate']):
                            eE_.attrib.clear()
                            eE_.text=str(ret_not_null(used_dict['variabilityestimate']))
                    else:
                        if pd.notna(used_dict['variabilityestimate']):
                            eE_=etree.SubElement(eE,'dispersionValue')
                            eE_.text=str(ret_not_null(used_dict['variabilityestimate']))

    def endpoint(self, element, indf, groups=None,sets=None):
        df_rename(indf)
        if groups is not None:
            groups = [i.lower() for i in groups]
        if sets is not None:
            sets=[i.lower() for i in sets]
        temp_dict = indf.loc[0, :].to_dict()
        if not pd.isna(temp_dict['countable']):
            element.xpath('countable')[0].text = str(temp_dict['countable']).lower()
        else:
            element.xpath('countable')[0].text = 'false'
        _et = element.xpath('unit')[0]
        if 'unit' in temp_dict.keys():
            if not pd.isna(temp_dict['unit']):
                _et.text = str(temp_dict['unit'])
        if str(temp_dict['countable']).lower()=='true':
            all_cats=set(indf['categoryid'])
            for cu_cat in element.xpath('categories/category'):
                if cu_cat.attrib['id'] not in all_cats:
                    element.xpath('categories')[0].remove(cu_cat)

        arm_groups=element.xpath('armReportingGroups/armReportingGroup')

        if arm_groups:
            for ord, et in enumerate(arm_groups):
                #self.add_report_value(et, str(temp_dict['countable']), indf, groups[ord])
                self.add_report_value2(et,indf,groups[ord])

        set_groups=element.xpath('subjectAnalysisSetReportingGroups/subjectAnalysisSetReportingGroup')
        if set_groups:
            for ord,et in enumerate(set_groups):
                self.add_report_value2(et,indf,sets[ord])

    def add_endpoints(self, indf, groups=None,sets=None, anaindf=None):
        df_rename(indf)
        if groups is not None:
            groups = [i.lower() for i in groups]
        if sets is not None:
            sets = [i.lower() for i in sets]
        endpoints = self.root.xpath("endPoints/endPoint")
        for et_ord, endpt in enumerate(endpoints):
            cindf = indf[indf['order'] == et_ord + 1]
            self.endpoint(endpt, cindf, groups,sets)
            if anaindf is not None:
                canaindf = anaindf[anaindf['order'] == et_ord + 1]
                if canaindf.shape[0] > 0:
                    self.add_statistical_analysis(endpt, canaindf)

    def add_country_count_exist(self, indf):
        df_rename(indf)
        country_counts = self.root.xpath('trialInformation/countrySubjectCounts/countrySubjectCount')

        indf_dict = {str(indf.loc[i,'eutctid']):indf.loc[i,'count'] for i in range(indf.shape[0])}
        for country in country_counts:
            if str(country.xpath('country/eutctId')[0].text) in indf_dict.keys():
                country.xpath('subjects')[0].text=str(
                    ret_not_null(indf_dict[str(country.xpath('country/eutctId')[0].text)]))
            else:
                country.xpath('subjects')[0].text='0'

    def add_country_count(self, indf,inexcel,sheet='country'):
        df_rename(indf)
        country_code=pd.read_excel(inexcel,sheet)
        df_rename(country_code)
        indf=indf.merge(country_code,on='abv3',how='left')
        indf['eutctid']=indf['identifier']
        country_counts = self.root.xpath('trialInformation/countrySubjectCounts')[0]
        for et in country_counts.getchildren():
            country_counts.remove(et)
        indf_dict = indf.to_dict(orient='records')
        for country in indf_dict:
            et = etree.SubElement(country_counts, 'countrySubjectCount')
            etree.SubElement(et, 'subjects').text = str(ret_not_null(country['count']))
            ct = etree.SubElement(et, 'country')
            euid = etree.SubElement(ct, 'eutctId')
            euid.text = str(ret_not_null(country['eutctid']))
            etree.SubElement(ct, 'version').text = str(ret_not_null(country['version']))


    def tofile(self, filename='ema.xml', pretty=True):
        if pretty:
            prettyXml(self.root, '\t', '\n')
        obj_xml = etree.tostring(self.root,
                                 pretty_print=True,
                                 xml_declaration=False)

        try:
            with open(filename, "w") as xml_writer:
                xml_writer.write("<?xml version='1.0' encoding='ASCII'?>\n")
                xml_writer.write(bytes.decode(obj_xml, 'UTF-8'))
        except IOError:
            pass


if __name__ == '__main__':
    test = postxml_ema(r'C:\Studies\posting\pilot\18933\18933_PharmaCM ERF_without value_2021-05-19.xml')
    test.addstudyno('18933')
    country = pd.read_excel(r'C:\Studies\posting\pilot\16244\posting_new.xlsx', 'COUNTRY', dtype='object')
    test.add_country_count(country, r'C:\Studies\posting\euctid_country.xlsx', 'country')

    AGE = pd.read_excel(r'C:\Studies\posting\pilot\16244\posting_new.xlsx', 'AGE', dtype='object')
    test.add_agegroup(AGE, '_cptog1')
    disp = pd.read_excel(r'C:\Studies\posting\pilot\16244\posting_new.xlsx', 'DISP', dtype='object')
    test.add_patientflow(disp, ['_COL1_', '_COL2_'])
    age_n = pd.read_excel(r'C:\Studies\posting\pilot\16244\posting_new.xlsx', 'AGE_N', dtype='object')
    test.add_age_cont_characters(age_n, ['col1', 'col2'])

    gender = pd.read_excel(r'C:\Studies\posting\pilot\16244\posting_new.xlsx', 'GENDER', dtype='object')
    test.add_gender_cat_characters(gender, ['col1', 'col2'], 'col3')

    demo_c = pd.read_excel(r'C:\Studies\posting\pilot\16244\posting_new.xlsx', 'DEMO_C', dtype='object')
    test.add_baseline_cat_characters(demo_c, ['col1', 'col2'], 'col3')

    demo_n = pd.read_excel(r'C:\Studies\posting\pilot\16244\posting_new.xlsx', 'DEMO_N', dtype='object')
    test.add_baseline_cont_characters(demo_n, ['col1', 'col2'])

    enpoints = pd.read_excel(r'C:\Studies\posting\pilot\18933\posting.xlsx', 'ENDPOINTS')
    stats = pd.read_excel(r'C:\Studies\posting\pilot\18933\posting.xlsx', 'STATS', dtype='object')
    test.add_endpoints(enpoints, ['col1', 'col2'], [], stats)
    test.tofile(r'C:\Studies\posting\pilot\18933\test_new.xml')

    test = postxml_ema(r'C:\Studies\posting\pilot\19773\19773_EudraCT Results Outline_No results value_2021-04-26.xml')
    enpoints = pd.read_excel(r'C:\Studies\posting\pilot\19773\testing.xlsx', 'test')
    test.add_endpoints(enpoints, ['group0'], ['set2','set3'])
    test.tofile(r'C:\Studies\posting\pilot\19773\test.xml',pretty=False)

    prettyxmlfile(r'C:\Studies\posting\pilot\18933\check.xml')











