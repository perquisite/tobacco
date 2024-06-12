# -*- coding:utf-8 -*-
# @ModuleName: over_all_description
# @Function: 
# @Author: huhonghui
# @email: 1241328737@qq.com
# @Time: 2021/10/13 17:57

# over_all_info data_item :([file, ContractCheckResult])
# ContractCheckResult(contract_type, factors, factors_ok, factors_error, factors_to_inform)
import xlwt

from ContractCheckResult import ContractCheckResult
from ContractType import ContractType
import time

type_list = {
    ContractType.BuySellContract: "标准-买卖合同",
    ContractType.ConstructionContract: "标准-建设工程施工合同",
    ContractType.RentContract: "标准-租赁合同",
    ContractType.PurchaseAndWarehousingContract: "标准-采购与服务供应商入库合同",
    ContractType.PropertyManagementContract: "标准-物业管理合同",
    ContractType.NotSure: "非标准-合同",
}
Remark_dict_risk = {
    # 买卖合同部分
    '一次性付款': ['不利于买方',
              '一次性支付全部货款将增加买方风险,建议修改为经双方协商一致，选择按照分期付款的方式进行付款：\n1、合同签订后【】个工作日内向卖方支付合同总额的【 】%(即【】元)作为预付款。\n2、所有货物交付、安装调试完成并通过验收取得验收合格文件后【】个工作日内，买方向卖方支付合同总额的【 】%(即【】元)；买方仅根据卖方提供的《送货单》对进场货物的数量进行确认，该确认不能免除卖方对所供货物应承担的质量保证责任（包括但不限于质量、规格、型号符合法律法规规定及合同约定）。卖方部分货物未交付完成的，买方该款约定的全部款项支付义务顺延至全部货物交付完毕。\n3、剩余【】%(即【】元)款项作为质保金，在交付完成后且质保期届满后，若无任何质量问题买方在【】个工作日内无息支付。'],
    '分期付款': ['(全局)不利于买卖双方',
             '分期付款期间出卖人解除合同的，可以向买受人请求支付该标的物的使用费\n分期付款期间如果买方逾期付款应该向卖方交滞纳金\n在分期付款的买卖合同中，买卖双方应注意关于“五分之一”的特别规定'],
    '发票': ['(全局)不利于买卖双方', '发票是保障买方售后和其他相关权益的重要凭证，是强有力的法定证据，尽可能完善发票条款，尽量避免法律风险。建议包括发票类型、税率，时间和次数以及未按约定开具发票的法律责任。'],
    '分批次交付': ['不利于买方', '明确分批次交付的时间节点，防止卖方逾期交货，如卖方应于【】年【】月【】日前将第【】批货物送至双方约定的交货地点。'],
    '货物符合相关标准': ['不利于买方',
                 '明确的质量标准有利于避免卖方交付不符合要求的货物，如(1)卖方保证提供的货物是全新、未曾使用过的，拥有合格证的合格货物，其名称、型号、规格、数量与本合同及附件约定相符，并符合相关法律、法规、行业规范、行业惯例等标准；(2)货物包装、外观符合行业相关标准；(3)货物涉及的知识产权等没有瑕疵；(4)货物相关文件齐备（系统货物出厂合格证、保修卡、说明书、操作手册等）'],
    '货物验收': ['(全局)不利于买卖双方',
             '验收时间的长短对买卖双方的权益均存在重要影响，建议约定为(1)如为买方，建议约定为：买方应在卖方交货/安装完毕后及时对货物进行验收，并在验收文件上签字确认,买方签字确认后视为卖方交付完成。(2)如为卖方，建议约定为：买方应在卖方交货/安装完毕后【】日内完成货物验收，验收通过的，买方在验收文件上签字确认,签字确认后视为卖方交付完成。如买方在卖上述期限内未签署验收文件，则视为卖方交付完成。'],
    '保质期': ['不利于买方',
            '保质期（质保期、保修期）期间卖方责任需明确约定，防止卖方怠于履行售后服务。建议约定为(1)保质期自验收合格之日起【】个(月/年），货物自身存在保质期的，以货物自身的保质期为准。(2)保质期内，因非人为原因发生故障、破损等质量问题造成不能正常使用的，卖方负责免费维修；对买方使用不当或其他因素造成货品不能正常使用的，卖方维修时需向买方收取维修成本费用。在接到买方维修要求后【】小时内做出回应，【】小时内安排维修人员达到产品现场，【】小时内完成修复；卖方承诺如果在【】小时内无法修复，将在【】小时内提供备用产品或与买方协商延长修复期间，以保证买方日常工作不受影响。(3)保质期内卖方未按约定时间上门维修，视为授权买方组织维修，并由卖方承担所有费用和质量责任，同时加收维修费用【】%的劳务费。(4)当发生货物损坏返修的情况时，该货物的保质期将重新计算。(5)货品保质期满后，卖方仍应在收到买方通知后进行维护服务，可收取相关的成本费用，具体费用由买卖双方协商确定。'],
    '逾期交付违约责任': ['不利于买方',
                 '逾期交付天数对应的违约金等需明确约定，防止责任不明，举证困难。建议修改为：\n1、如为卖方，建议修改为：因买方过错导致逾期支付合同价款的，买方应按逾期付款部分【】‰/天的比例向卖方支付违约金；买方逾期支付合同价款超过【】日的，卖方有权选择单方面解除本合同，此时双方按照已经交付完成的货物进行结算，买方应按合同总金额的【】％向卖方支付违约金。\n2、如为买方，建议修改为：因买方过错导致逾期支付合同价款的，买方应按逾期付款部分【】‰/天的比例向卖方支付违约金；买方逾期支付合同价款超过【】日的，卖方有权选择单方面解除本合同，此时双方按照已经交付完成的货物进行结算，买方应按逾期付款部分的【】％向卖方支付违约金。'],
    '单方解除合同违约责任': ['(全局)不利于买卖双方',
                   '违约责任需明确约定，防止责任不明，举证困难。建议修改为：除本合同另有约定外，任何一方均无权擅自单方解除合同。若任何一方有违法单方解除合同行为的，无论解除行为是否生效，均须向另一方支付违约金。违约金为本合同总金额的【】%。'],
    '侵犯第三方权利违约责任': ['不利于买方',
                    '违约责任需明确约定，防止责任不明，举证困难。建议修改为：卖方保证所提供的货物或其任何一个组成部分均不会侵犯任何第三方的专利权、商标权、著作权、商业秘密等合法权利；如出现侵权情形，卖方应承担违约责任。买方有权解除本合同，卖方应退还全部已收款项，并按合同总金额的【】％向买方支付违约金。'],
    # 物业管理部分
    '安全保卫': ['不利于：甲方',
             '安全保卫要求及标准应包括：\n1、秩序维护与安全防范落实到位：\n（1）保安人员持证上岗；\n（2）保安人员实行24小时值班及巡查制度；熟悉小区的环境，文明礼貌，工作负责；有交接班记录和值班巡查记录，有来访登记制度；\n（3）夜间对服务范围内重点部位、道路进行防范检查和巡视，做好记录；\n（4）危及人身安全处有明显标识和具体的防范措施；\n（5）门禁、监控等安全防范设施正常运行、维护及时，有定期维护记录。\n2、消防管理制度健全：\n（1）消防设备设施完好无损，可随时启用；消防通道畅通；\n（2）消防管理制度健全，有消防应急预案；每年进行一次消防演练及一次以上消防培训；\n（3）消防设施有明显标志，定期对消防设施进行巡视、检查和维护，并有记录。\n3、停车场及车辆管理有序：\n（1）机动车停车场管理制度完善，车辆停放有序，进出有登记；\n（2）非机动车车辆管理制度完善，按规定位置停放，管理有序，未占道停车；\n（3）主要道路及停车场有交通标志；\n（4）交通设施（道闸、交通标识等）能正常使用；地下停车场照明、给排水、通风系统正常运行，各类标识清晰；\n（5）停车收费严格按规定执行，无乱收费现象。'],
    '卫生保洁': ['不利于：甲方',
             '卫生保洁要求及标准应包括：\n1、环卫设备管理到位与公共环境保持良好：\n（1）生活垃圾封闭式管理，设有垃圾箱、果皮箱；垃圾日产日清；\n（2）房屋共用部位、共用设施设备无蚁害；\n（3）房屋共用部位保持清洁，无乱贴、乱画，无擅自占用和堆放杂物现象；楼梯扶栏、天台、公共玻璃窗等保持洁净；\n（4）商业网点管理有序；无乱设摊点、广告牌和乱贴、乱画现象；\n（5）无违反规定饲养宠物、家禽、家畜；\n（6）排放油烟、噪音等符合国家标准，外墙无污染。\n2、清洁服务落实到位：\n（1）清洁卫生实行责任制，有专职的清洁人员并明确责任范围，实行标准化保洁。定期进行卫生消毒灭杀；\n（2）小区内道路等共用场地和绿化带内无纸屑、烟头等废弃物和垃圾；\n（3）在雨、雪天气及时对小区内主干路和屋（棚）顶积水、积雪进行清扫。'],
    '接待及会务服务': ['不利于：甲方',
                '接待及会务服务要求及标准应包括：\n（1）建立24小时值班制度，有执行值班制度的情况汇报及相关记录；\n（2）有单独的投诉回访制度，投诉回访记录规范齐全；\n（3）每年定期向住（用）户发放物业服务工作征求意见单，并对意见及时整理和处理，对合理的采纳及时整改，满意率达85%以上。'],
    '员工食堂': ['不利于：甲方',
             '员工食堂服务的要求及标准应重点包括：\n1、周边环境整洁；\n（1）门前垃圾必须及时清扫，保持店面周边无杂物、无垃圾、无积水、无污物。\n（2）设置专门餐厨垃圾投放容器（配有盖子）。\n（3）及时清理垃圾桶，不留异物、不产生异味，不对周边卫生和空气造成污染。\n2、就餐场所干净；\n（1）就餐区日常清洁：指定专门卫生人员负责就餐区日常清洁，保持桌面、座椅、墙面、地面清洁，门窗洁净明亮。\n（2）定期清洁就餐场所的空调、排风扇、地毯等设施或物品。\n（3）配备洗手液（皂）、消毒液、擦手纸、干手器等。\n（4）就餐场所无老鼠、蟑螂、苍蝇等。\n3、后厨合规达标：\n（1）每天对后厨进行全面清洗，保持后厨清洁卫生。并定期定时对后厨设施设备进行消毒。\n（2）及时清理餐厨废弃物、污物、垃圾等，保持后厨地面整洁、排水沟通畅，不留异物、不产生异味。\n（3）加强排油烟、排气、通风等设施设备的清理、维修、保养，确保设施设备满足卫生要求，保持设施处于正常工作状态。\n（4）加强防蝇、防鼠、防病虫害管理，定期开展灭蝇、除鼠、杀虫。\n（5）做到无蜘蛛网、无积尘、无虫害、无鼠迹，确保食物、设施设备不受污染。\n4、餐饮用具洁净\n（1）落实餐饮具清洗消毒制度，配备清洗消毒设施设备、明确岗位职责。\n（2）确保餐用具严格清洗消毒后使用，餐饮用具清洗消毒不留残渣、不积水、不油滑，达到“光、洁、涩、干”效果。\n（3）加强消毒后的保洁管理，消毒后的餐饮具入柜、密闭、不外露，防止消毒后的餐饮具被重新污染。\n（4）加强对集中消毒的餐饮具采购、验收管理，索取餐饮具集中消毒企业的合法资质证照，索取每批次餐饮具的消毒合格证明文件。\n5、从业人员健康\n（1）建立健全从业人员健康档案。从事接触直接入口食品的工作人员必须每年参加体检，取得健康证明后方可上岗。凡患有有碍食品安全疾病并从事直接入口食品的工作人员，一律不得上岗。\n（2）从业人员进入厨房时，必须更衣、戴帽、洗手。从业人员上岗时佩戴饰物不得外露。\n（3）厨房区域内严禁吸烟。\n（4）从业人员在从事任何可能污染双手的活动后必须洗手。'],
    '设施设备维修': ['不利于：甲方',
               '设施设备维修的要求及标准应重点包括：\n1、房屋维护计划落实到位：\n①制定年度房屋维修养护计划并有实施记录；\n②维修及时，临修急修及时率100%，返修率不高于2%，并有回访记录；\n③对房屋共用部位进行日常管理和维修养护，有维修记录和保养记录；\n④定期巡视（每月不少于二次）房屋共用部位的地、墙等小区房屋单元门、楼梯通道以及其他共用部位的门窗、玻璃等，做好巡查记录，并及时维修养护；\n2、公共设备设施日常管理和养护到位：\n（1）制定设施设备维修养护计划并实施。有操作规程与维保记录，无安全隐患；\n（2）建立共用设施设备清册档案（设备台帐），有设备的运行、检查、保养、维修记录；有设备标识；\n（3）对共用设施设备适时组织巡查（按操作规程巡查），有巡查记录，小修范围及时修复；大中修或需更新改造的，提出报告与建议，并有实施记录；\n3、排水、排污、道路设施通畅\n（1）排水、排污管道、化粪池完整通畅，无堵塞外溢现象；\n（2）公共区域内的雨水、污水管每半年检查、疏通一次，雨水、污水井每半年检查、清掏一次；化粪池清掏每年一次，每季检查一次，并有相关记录；\n（3）道路通畅，路面平整；井盖无缺损、无丢失，场地（景观、健身器材、建筑小品等）维护良好。\n4、供水、供电设备管理规范：\n（1）供水设备运行正常，设施完好、无渗漏、无污染；\n（2）二次生活用水有严格的保障措施，按规定清洗水池，水质符合卫生标准，取得卫生许可证，有专人负责，操作人员取得健康体检合格证；\n（3）制订停水停电及事故处理方案。停水停电应及时告知业主；\n（4）制订供电系统管理措施并严格执行，记录完整；专变供电设备运行正常，配电室管理符合规定，路灯、楼道灯等公共照明设备完好。\n5、电梯管理制度健全，管理规范：\n（1）电梯由专业维保公司进行维修保养，年检标识有效；轿箱、井道和电梯机房干净整洁，通风、照明良好；\n（2）日常维修、保养人员持证上岗，保养维修记录齐全；无安全事故；\n（3）电梯应急对讲畅通，制订出现故障的应急处理方案；在电梯轿厢内张贴安全注意事项和24小时应急联系电话，设置紧急装置。'],
    '绿化养护': ['不利于：甲方',
             '绿化养护的要求及标准应包括：\n（1）绿化带无明显裸露土地；树木生长正常，无死树和明显枯枝死枝，树木无明显钉栓、捆绑现象；\n（2）绿地无改变用途，无损坏、践踏、占用现象；\n（3）绿化有专人养护管理，花草树木长势良好，修剪整齐美观；\n（4）及时灭治虫害，有灭杀记录，无枯死花草、树木。'],
    '履约保证金退还': ['不利于：甲方',
                '建议约定为：\n项目期限届满，在乙方按要求提交《履约保证金退还申请》之日起【】日内，甲方根据乙方的考核结果对履约保证金进行扣减，上述期限届满后【】日内，甲方根据扣减结果向乙方无息退还履约保证金。'],
    '检查考核': ['不利于：甲方', '风险提示：为避免争议，检查考核需具有完备的考核依据，同时需在合同中明确考核的主体部门或岗位。'],
    '乙方服务不符合约定的违约责任': ['不利于：甲方',
                       '建议修改为：\n乙方违反本合同的约定，未能达到所约定的管理目标，甲方有权要求乙方限期整改并达到本合同约定；逾期未整改的，或整改后仍不符合合同约定的，甲方可以终止本合同，乙方应向甲方支付【】元的违约金，该违约金不足赔偿甲方损失的，乙方还应补足。'],
    # 采购部分
    '甲方提供技术需求及其他资料文件': ['不利于：甲方', '建议更改为：\n'
                                      '乙方在接受具体项目后，甲方可根据实际情况及自身需求提供能够全面反映采购预期目标的技术需求及其他相关资料或文件。'],
    '乙方产品与服务标准': ['不利于：甲方', '在订立具体采购合同时，建议明确特定产品应符合的标准，'
                            '防止约定不明，难以追究乙方未按约定提供产品和服务的违约责任。'],
    '乙方违约责任': ['不利于：甲方', '风险提示：违约责任的量化有利于追究乙方违约责任，请根据示范文本进行完善。'],
    '维保期（保质期（质保期、保修期）': ['不利于：甲方', '建议约定为：\n1、保质期自验收合格之日起【】个(月/年），货物自身存在保质期的，以货物自身的保质期为准。\n'
                                   '2.保质期内，因非人为原因发生故障、破损等质量问题造成不能正常使用的，乙方负责免费维修；对甲方使用'
                                   '不当或其他因素造成货品不能正常使用的，乙方维修时需向甲方收取维修成本费用。在接到甲方维修要求后【】'
                                   '小时内做出回应，【】小时内安排维修人员达到产品现场，【】小时内完成修复；乙方承诺如果在【】小时内无法'
                                   '修复，将在【】小时内提供备用产品或与甲方协商延长修复期间，以保证甲方日常工作不受影响。\n'
                                   '3.保质期内乙方未按约定时间上门维修，视为授权甲方组织维修，并由乙方承担所有费用和质量责任，同时加收维修费用【】%的劳务费。\n'
                                   '4.当发生货物损坏返修的情况时，该货物的保质期将重新计算。\n'
                                   '5.货品保质期满后，乙方仍应在收到甲方通知后进行维护服务，可收取相关的成本费用，具体费用由双方协商确定。'],
    '供应商关于取消入库资格的规定': ['不利于：双方', '供应商关于取消入库资格的规定是甲方执行取消乙方入库资格的重要依据，故建议作为本合同附件。'],
    # 租赁合同部分
    '出租方': ['不利于：承租方',
                '主体资格的审查风险提示：应对出租人的资格进行审查，审核其是否具有房屋的处分权。查验出租人的房屋产权凭证，如果出租人非为原房东，需其提供与原房东签订的租赁协议，查看其中关于的转租的条款约定，最好要求原房东到场一起签署。'],
    '抵押': ['不利于：承租方', '房屋权属的审查风险提示：请核实房屋是否有抵押等其他权利，防止房屋因第三方原因导致无法租赁。'],
    '租赁费用': ['(全局)不利于：双方',
             '租金支付方式的确定风险提示：双方应根据不同的主体地位明确租金缴纳方式：（1）如为承租方，为尽可能降低支付压力，建议约定按月交纳（如按季度、年缴纳存在优惠，请综合考虑后决定）。（2）如为出租方，建议约定明确的缴纳期限及逾期支付租金的滞纳金。'],
    '租赁房屋交付': ['(全局)不利于：双方',
               '租赁房屋的交付及接收风险提示：在交付时应采用清单的形式，明确交付的租赁房屋中的全部设施设备；除此之外：（1）如为承租方，建议延长房屋接收时的验收期限，并明确验收未通过时出租方的责任。（2）如为出租方，建议尽可能缩短房屋接收时的验收期限，并明确承租方逾期验收/接收的违约责任。'],
    '返还时间': ['(全局)不利于：双方',
             '房屋返还时间风险提示：双方应根据不同的主体地位明确房屋返还时间：（1）如为承租方，建议尽可能延长返还时间。（2）如为出租方，建议尽可能缩短房屋返还时间，并明确延期返还的违约责任。'],
    '遗留物品处理': ['不利于：承租方', '遗留物品处理的方式风险提示：如为承租方，建议仅约定出租方具有留置权，避免财产遭受较大损失。'],
    '转租责任': ['不利于：承租方', '转租责任风险提示：（1）如为承租方，建议约定在一定条件下可对房屋进行转租。（2）如为出租方，建议约定较为严格的擅自转租的违约责任。'],
    '违约责任': ['(全局)不利于：双方',
             '违约责任风险提示：（1）逾期交付房屋、支付租金、返还房屋的天数对应的违约金等需明确约定，防止责任不明，举证困难；（2）逾期履行上述责任超过多少时间另一方可解除合同，要求对方承担违约责任需进行明确；（3）违约金数额需明确且保证合法合理。'],
}


def get_over_all_file(over_all_info, export_dir):
    need_to_write = [0] * len(over_all_info)

    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('合同审查结果总览')
    worksheet.col(0).width = 256 * 30
    worksheet.col(1).width = 256 * 30
    worksheet.col(2).width = 256 * 20
    worksheet.col(3).width = 256 * 30
    worksheet.col(4).width = 256 * 30
    worksheet.col(5).width = 256 * 40
    worksheet.col(6).width = 256 * 40
    worksheet.col(7).width = 256 * 40
    worksheet.col(8).width = 256 * 40
    worksheet.col(9).width = 256 * 40
    worksheet.write(0, 0, "合同文件", header_style())
    worksheet.write(0, 1, "路径", header_style())
    worksheet.write(0, 2, "合同类型", header_style())
    worksheet.write(0, 3, "提取要素数量", header_style())
    worksheet.write(0, 4, "要素错误数量", header_style())
    worksheet.write(0, 5, "要素需要相关部门确认的数量", header_style())
    worksheet.write(0, 6, "非标准合同-完整性检查数量", header_style())
    worksheet.write(0, 7, "非标准合同-完整性缺失数量", header_style())
    worksheet.write(0, 8, "非标准合同-风险性检查数量", header_style())
    worksheet.write(0, 9, "非标准合同-风险性警示数量", header_style())

    for i in range(len(over_all_info)):
        if isinstance(over_all_info[i], list):
            file_path = over_all_info[i][0]  # such as C:/Users/12413/Desktop/买卖合同全错版1(已审查).docx
            file = file_path.split("/")[-1]
            contractCheckResult = over_all_info[i][1]
            contract_type = contractCheckResult.type
            if contract_type == ContractType.NotSure:
                worksheet.write(i + 1, 0, file)
                worksheet.write(i + 1, 1, file_path)
                worksheet.write(i + 1, 2, type_list[contract_type], center())
                worksheet.write(i + 1, 3, "-")
                worksheet.write(i + 1, 4, "-")
                worksheet.write(i + 1, 5, "-")

                # 非标准合同信息统计:完整性检查的总项目数、缺失项
                integrity_check_num = 0
                integrity_find_num = 0
                for item in contractCheckResult.factors_error.values():
                    integrity_check_num += len(item["check_lsit"])
                    integrity_find_num += sum(item["check_lsit"])
                worksheet.write(i + 1, 6, integrity_check_num)
                worksheet.write(i + 1, 7, integrity_check_num - integrity_find_num)

                # 非标准合同信息统计:风险性检查的总项目数、缺失项
                risk_check_num = len(contractCheckResult.factors_to_inform.keys())
                risk_find_num = 0
                for item in contractCheckResult.factors_to_inform.values():
                    risk_find_num += item["check_lsit"][0]
                worksheet.write(i + 1, 8, risk_check_num)
                worksheet.write(i + 1, 9, risk_check_num - risk_find_num)

                # 是否需要write明细
                if integrity_check_num - integrity_find_num != 0 or risk_check_num - risk_find_num != 0:
                    need_to_write[i] = 1

            else:
                factors_error_num = len(contractCheckResult.factors_error)
                factors_to_inform_num = len(contractCheckResult.factors_to_inform)
                factors_num = len(contractCheckResult.factors)
                worksheet.write(i + 1, 0, file)
                worksheet.write(i + 1, 1, file_path)
                worksheet.write(i + 1, 2, type_list[contract_type], center())
                worksheet.write(i + 1, 3, factors_num)
                worksheet.write(i + 1, 4, factors_error_num)
                worksheet.write(i + 1, 5, factors_to_inform_num)
                worksheet.write(i + 1, 6, "-")
                worksheet.write(i + 1, 7, "-")
                worksheet.write(i + 1, 8, "-")
                worksheet.write(i + 1, 9, "-")

                # 是否需要write明细
                if factors_error_num != 0 or factors_to_inform_num != 0:
                    need_to_write[i] = 1

        if isinstance(over_all_info[i], str):
            file_path = over_all_info[i][2:]  # such as C:/Users/12413/Desktop/买卖合同全错版1(已审查).docx
            file = file_path.split("/")[-1]
            if over_all_info[i][0:2] == "空白":
                worksheet.write(i + 1, 0, file)
                worksheet.write(i + 1, 1, file_path)
                worksheet.write(i + 1, 2, "空白文件")
                worksheet.write(i + 1, 3, "-")
                worksheet.write(i + 1, 4, "-")
                worksheet.write(i + 1, 5, "-")

            if over_all_info[i][0:2] == "未知":
                worksheet.write(i + 1, 0, file)
                worksheet.write(i + 1, 1, file_path)
                worksheet.write(i + 1, 2, "合同类型不在常用合同库中")
                worksheet.write(i + 1, 3, "-")
                worksheet.write(i + 1, 4, "-")
                worksheet.write(i + 1, 5, "-")

    for i in range(len(over_all_info)):
        if isinstance(over_all_info[i], list):
            file_path = over_all_info[i][0]
            file = file_path.split("/")[-1]
            contractCheckResult = over_all_info[i][1]
            contract_type = contractCheckResult.type
            # 非标准合同检查明细标签页
            if contract_type == ContractType.NotSure:
                if need_to_write[i] == 0:
                    continue
                worksheet = workbook.add_sheet(file.split(".")[0])

                worksheet.col(0).width = 256 * 42
                worksheet.col(1).width = 256 * 32
                worksheet.col(2).width = 256 * 56

                worksheet.write(0, 0, "批注提示项", header_style())
                worksheet.write(0, 1, "问题类型", header_style())
                worksheet.write(0, 2, "描述", header_style())

                contractCheckResult = over_all_info[i][1]
                factors = contractCheckResult.factors
                factors_error = contractCheckResult.factors_error
                factors_to_inform = contractCheckResult.factors_to_inform
                line = 1
                for key in factors.keys():
                    for item in factors[key]:
                        worksheet.write(line, 0, key)
                        worksheet.write(line, 1, "完整性检查-缺失要素")
                        worksheet.write(line, 2, item)
                        line += 1
                for key in factors_to_inform:
                    info = factors_to_inform[key]
                    if info["check_flg"]:
                        worksheet.write(line, 0, key)
                        worksheet.write(line, 1, "风险性检查-风险警示")
                        worksheet.write(line, 2, Remark_dict_risk[key][0] + "\n" + Remark_dict_risk[key][1])
                        line += 1


            # 标准合同
            else:
                if len(over_all_info[i][1].factors_error) == 0 and len(over_all_info[i][1].factors_to_inform) == 0:
                    continue

                worksheet = workbook.add_sheet(file.split(".")[0])

                worksheet.col(0).width = 256 * 42
                worksheet.col(1).width = 256 * 32
                worksheet.col(2).width = 256 * 56

                worksheet.write(0, 0, "存在问题的要素名称", header_style())
                worksheet.write(0, 1, "问题级别", header_style())
                worksheet.write(0, 2, "描述", header_style())

                contractCheckResult = over_all_info[i][1]
                factors_error = contractCheckResult.factors_error
                factors_to_inform = contractCheckResult.factors_to_inform

                index = 0

                for key, value in factors_error.items():
                    index = index + 1
                    worksheet.write(index, 0, key)
                    worksheet.write(index, 1, "填写错误")
                    worksheet.write(index, 2, value)

                for key, value in factors_to_inform.items():
                    index = index + 1
                    worksheet.write(index, 0, key)
                    worksheet.write(index, 1, "内容需要相关部门审核确认")
                    worksheet.write(index, 2, value)

    # name = time.strftime("%Y-%m-%d-%H-%M", time.localtime())

    workbook.save(export_dir + "/" + "审查总览.xls")
    # workbook.close()


def header_style():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.bold = True
    font.height = 20 * 14
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 0x0D
    ali = xlwt.Alignment()
    ali.horz = 0x02
    style.font = font
    style.pattern = pattern
    style.alignment = ali
    return style


def center():
    style = xlwt.XFStyle()
    ali = xlwt.Alignment()
    ali.horz = 0x02
    style.alignment = ali
    return style


if __name__ == '__main__':
    over = []
    contract_type = ContractType.NotSure
    s = "{'合同主体信息': ['甲方住所地', '甲方统一社会信用代码/身份证号码', '甲方联系电话', '甲方电子邮件', '乙方法定代表人/负责人', '乙方住所地', '乙方统一社会信用代码/身份证号码', '乙方联系电话', '乙方电子邮箱'], '货物（标的）信息': ['规格型号', '计量单位', '总价'], '货物（标的）包装信息': ['包装材料', '包装总价'], '收货联系人信息': ['收货人姓名', '收货人联系方式'], '货物（标的）运输信息': ['运输时间', '运输方式', '运费负担', '运费保险费用承担', '运输通知'], '货物（标的）交付方式信息': ['交付方式'], '货物（标的）交付完成条件信息': ['交付完成条件'],'货物（标的）风险转移信息': ['风险转移'], '货物（标的）安装信息': ['安装时间', '安装地点', '安装费用承担'], '开票信息': ['开票类型', '开票公司名称', '纳税识别号', '地址', '电话'], '收款账户信息': ['收款开户名', '收款开户行'], '知识产权条款': ['知识产权条款']}"
    factors = eval(s)
    factors_ok = []
    # 完整性字典
    s = "{'合同主体信息': {'check_flg': False, 'flg_str': '（2）乙方交货的品种、数量、型号等若与合同约定不符，乙方应在7个工作日内无条件更换，并赔偿由此给甲方带来的损失，且不免除乙方延期交付的违约责任', 'check_lsit': [1, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0]}, '货物（标的）信息': {'check_flg': False, 'flg_str': '1、货物名称、型号规格、单位、数量、单价、金额及合同价', 'check_lsit': [1, 0, 0, 1, 1, 1, 1, 0]}, '货物（标的）包装信息': {'check_flg': False, 'flg_str': '4、质量标准：乙方保证提供的产品系全新行货产品，包装完整，配件齐全，质量完好，符合国家3C认证（中国强制性产品认证制度）标准', 'check_lsit': [1, 0, 0]}, '收货联系人信息': {'check_flg': False, 'flg_str': '', 'check_lsit': [0, 0]}, '货物（标的）运输信息': {'check_flg': False, 'flg_str': '', 'check_lsit': [0, 1, 0, 0, 0, 0]}, '货物（标的）交付方式信息': {'check_flg': False, 'flg_str': '', 'check_lsit': [0]}, '货物（标的）交付地点信息': {'check_flg': True, 'flg_str': '2、交货地点和方式：崇州市崇阳街道蜀州中路341号崇州市烟草专卖局（分公司）', 'check_lsit': [1]}, '货物（标的）验收标准信息': {'check_flg': True, 'flg_str': '6、付款方式：货物验收合格，据实结算，在收到乙方开具合法有效全额增值税专用发票之日起30个工作日内一次性转账付清全部货款', 'check_lsit': [1]}, '货物（标的）交付完成条件信息': {'check_flg': False, 'flg_str': '', 'check_lsit': [0]}, '货物（标的）风险转移信息': {'check_flg': False, 'flg_str': '', 'check_lsit': [0]}, '货物（标的）安装信息': {'check_flg': False, 'flg_str': '', 'check_lsit': [0, 0, 0]}, '付款方式信息': {'check_flg': True, 'flg_str': '6、付款方式：货物验收合格，据实结算，在收到乙方开具合法有效全额增值税专用发票之日起30个工作日内一次性转账付清全部货款', 'check_lsit': [1]}, '开票信息': {'check_flg': False, 'flg_str': '银行账号： 51001416177051508447', 'check_lsit': [0, 0, 0, 0, 0, 1, 1]}, '收款账户信息': {'check_flg': False, 'flg_str': '银行账号： 51001416177051508447', 'check_lsit': [0, 0, 1]}, '质量保证责任信息': {'check_flg': True, 'flg_str': '非产品质量问题（人为损坏、使用不当或不可抗拒之力）造成的产品损坏不在保修范围内', 'check_lsit': [1]}, '违约责任条款约定': {'check_flg': True, 'flg_str': '（1）逾期交货或付款的，每天按合同总价的千分之三向对方支付违约金，逾期超过30日，守约方可单方解除合同并向违约方追索损失', 'check_lsit': [1]}, '送达通知条款': {'check_flg': True, 'flg_str': '每三个月进行一次远程售后回访，将所反馈问题2小时内做出回应，保证72小时内解决问题', 'check_lsit':[1]}, '保密条款': {'check_flg': True, 'flg_str': '若更换后仍与合同约定不符，乙方应向甲方支付合同总价30%的违约金，且甲方可单方解除合同', 'check_lsit': [1]}, '知识产权条款': {'check_flg': False, 'flg_str': '', 'check_lsit': [0]}, '不可抗力条款': {'check_flg': True, 'flg_str': '8、违约责任：双方应认真履行本合同中规定的内容，任何一方不得擅自变更，否则需承担相应的违约责任', 'check_lsit': [1]}, '争议解决条款': {'check_flg': True, 'flg_str': '9、争议解决方式：本合同在履行过程中，发生的争议由双方协商解决，在双方协商不能达成一致的情况下，任何一方都可以向崇州市人民法院提起诉讼', 'check_lsit': [1]}, '合同生效条款': {'check_flg': True, 'flg_str': '10、本合同一式肆份，甲方执叁份，乙方执壹份，自双方授权代表签字并加盖双方公章或合同专用章后生效，具有同等法律效力', 'check_lsit': [1]}, '合同变更条款': {'check_flg': True, 'flg_str': '8、违约责任：双方应认真履行本合同中规定的内容，任何一方不得擅自变更，否则需承担相应的违约责任', 'check_lsit': [1]}}"
    factors_error = eval(s)
    # 风险性字典
    s = "{'一次性付款': {'check_flg': True, 'flg_str': '6、付款方式：货物验收合格，据实结算，在收到乙方开具合法有效全额增值税专用发票之日起30个工作日内一次性转账付清全部货款', 'check_lsit': [1]}, '分期付款': {'check_flg': False, 'flg_str': '', 'check_lsit': [0]}, '发票': {'check_flg': True, 'flg_str': '6、付款方式：货物验收合格，据实结算，在收到乙方开具合法有效全额增值税专用发票之日起30个工作日内一次性转账付清全部货款', 'check_lsit': [1]}, '分批次交付': {'check_flg': False, 'flg_str': '', 'check_lsit': [0]}, '货物符合相关标准': {'check_flg': True, 'flg_str': '4、质量标准：乙方保证提供的产品系全新行货产品，包装完整，配件齐全，质量完好，符合国家3C认证（中国强制性产品认证制度）标准', 'check_lsit': [1]}, '货物验收': {'check_flg': True, 'flg_str': '3、交货时间：合同签订后20个工作日内交货并通过验收', 'check_lsit': [1]}, '保质期': {'check_flg': True, 'flg_str': '在质保期外,提供设备的更换、维修只收取零配件成本费用,不收取人工技术和服务费用，免费提供应急包的使用培训', 'check_lsit': [1]}, '逾期交付违约责任': {'check_flg': True, 'flg_str': '（1）逾期交货或付款的，每天按合同总价的千分之三向对方支付违约金，逾期超过30日，守约方可单方解除合同并向违约方追索损失', 'check_lsit': [1]}, '单方解除合同违约责任': {'check_flg': False, 'flg_str': '', 'check_lsit': [0]}, '侵犯第三方权利违约责任': {'check_flg': False, 'flg_str': '', 'check_lsit': [0]}}"
    factors_to_inform = eval(s)
    c = ContractCheckResult(contract_type, factors, factors_ok, factors_error, factors_to_inform)
    over.append(["D:/测试.doc", c])
    get_over_all_file(over, "D:")
