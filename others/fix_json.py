#!/usr/bin/env python

import sys
import json


def replace_json_value(order_json):
    key_list = [
        "streetFenceName", "encourageAutonomy", "feedbackReasonOrigin", "channel",
        "businessId", "ruleRequirementConfig", "curRejectReason", "source",
        "gridFenceId", "workFenceId", "feedbackReasonRecord", "placeCategory",
        "streetFenceId", "gridFenceName", "healthyStatus", "creatorSymbol",
        "workFenceName", "initialSource", "newInterface", "relateOrderRequirementTypeDesc",
        "orderGroupRegionId", "relateOrderRequirementType", "pageCode", "isCommonCreate",
        "isMoleculeOrder", "jobId", "ogComplaint", "jobDetailId", "identifyBikeState",
        "circleOrder", "orderGroupContainRegionIds", "dispatchOrderSource", "t0Created",
        "orderServiceOnline", "strategyScene", "recommendNumber"
    ]

    json_map = json.loads(order_json)
    replace_map = {}

    if ("meta_data_info" in json_map and
            "meta_data_keyword" in json_map["meta_data_info"]):
        second_map = json_map["meta_data_info"]["meta_data_keyword"]
        for key in key_list:
            if key in second_map and second_map[key]:
                replace_map[key] = second_map[key]

    if replace_map:
        if "job_data_info" not in json_map:
            json_map["job_data_info"] = {}
        if "meta_data_keyword" not in json_map["job_data_info"]:
            json_map["job_data_info"]["meta_data_keyword"] = {}
        json_map["job_data_info"]["meta_data_keyword"].update(replace_map)

    return json.dumps(json_map, ensure_ascii=False)


for line in sys.stdin:
    order_id = line.strip().split("\t")[0]
    order_json = line.strip().split("\t")[1]
    updated_json = replace_json_value(order_json)
    print(order_id + "\t" + updated_json)

# 455355013	{"meta_data_info":{"meta_data_keyword":{"gridFenceName":"和平B1","workFenceId":"56752068","streetFenceName":"南市街道","streetFenceId":"37752197","gridFenceId":"129783897","newInterface":"true","relateOrderRequirementTypeDesc":"特殊保障&日常要求","encourageAutonomy":"13","orderGroupRegionId":"022","channel":"find_bike","businessId":"2024092415291335630765_1_t0_independent_work","ruleRequirementConfig":"any","source":"findbike_launchApply_add_point","relateOrderRequirementType":"regularRequirement","isCommonCreate":"true","isMoleculeOrder":"true","jobId":"126328802","ogComplaint":"false","placeCategory":"system_recommendation","jobDetailId":"258879997","identifyBikeState":"3","circleOrder":"false","healthyStatus":"[\"1\",\"4\"]","orderGroupContainRegionIds":"[\"022\"]","creatorSymbol":"system_recommend_to_t0","workFenceName":"经营-商圈-印保利广场右下角-B1南市街-晚高峰","dispatchOrderSource":"findbike_launchApply_add_point","initialSource":"t0_independent_work"},"meta_data":{"newInterface":"true","encourageAutonomy":"13","businessId":"2024092415291335630765_1_t0_independent_work","channel":"find_bike","source":"findbike_launchApply_add_point","relateOrderRequirementType":"regularRequirement","isMoleculeOrder":"true","ogComplaint":"false","identifyBikeState":"3","circleOrder":"false","healthyStatus":"[\"1\",\"4\"]","creatorSymbol":"system_recommend_to_t0","dispatchOrderSource":"findbike_launchApply_add_point"}},"job_data_info":{"meta_data_keyword":{"newInterface":"true","relateOrderRequirementTypeDesc":"特殊保障&日常要求","encourageAutonomy":"13","orderGroupRegionId":"022","channel":"find_bike","businessId":"2024092415291335630765_1_t0_independent_work","ruleRequirementConfig":"any","source":"findbike_launchApply_add_point","relateOrderRequirementType":"regularRequirement","isCommonCreate":"true","isMoleculeOrder":"true","jobId":"126328802","workFenceId":"56752068","ogComplaint":"false","placeCategory":"system_recommendation","jobDetailId":"258879997","identifyBikeState":"3","circleOrder":"false","healthyStatus":"[\"1\",\"4\"]","orderGroupContainRegionIds":"[\"022\"]","creatorSymbol":"system_recommend_to_t0","workFenceName":"经营-商圈-印保利广场右下角-B1南市街-晚高峰","dispatchOrderSource":"findbike_launchApply_add_point","initialSource":"t0_independent_work","roadScopeFenceName":"","streetFenceName":"南市街道","streetFenceId":"37752197","gridFenceName":"和平B1","roadScopeFenceId":"","gridFenceId":"129783897","feedbackReasonOrigin":"[{\"reasons\":[{\"topic\":{\"code\":\"T100\",\"type\":\"TOPIC\",\"choices\":[{\"code\":\"106\",\"type\":\"CHOICE\",\"value\":\"交通拥堵无法到达\"}]},\"photos\":{},\"text\":{}}]}]","curRejectReason":"交通拥堵无法到达","feedbackReasonRecord":"[{\"reason\":[{\"code\":\"106\",\"value\":\"交通拥堵无法到达\"}],\"photo\":[]}]","pageCode":"UNLOAD#REFUSE_ORDER"},"meta_data":{"newInterface":"true","encourageAutonomy":"13","businessId":"2024092415291335630765_1_t0_independent_work","channel":"find_bike","source":"findbike_launchApply_add_point","relateOrderRequirementType":"regularRequirement","isMoleculeOrder":"true","ogComplaint":"false","identifyBikeState":"3","circleOrder":"false","healthyStatus":"[\"1\",\"4\"]","creatorSymbol":"system_recommend_to_t0","dispatchOrderSource":"findbike_launchApply_add_point","feedbackReasonOrigin":"[{\"reasons\":[{\"topic\":{\"code\":\"T100\",\"type\":\"TOPIC\",\"choices\":[{\"code\":\"106\",\"type\":\"CHOICE\",\"value\":\"交通拥堵无法到达\"}]},\"photos\":{},\"text\":{}}]}]","curRejectReason":"交通拥堵无法到达","pageCode":"UNLOAD#REFUSE_ORDER","workFenceId":"56752068","feedbackReasonRecord":"[{\"reason\":[{\"code\":\"106\",\"value\":\"交通拥堵无法到达\"}],\"photo\":[]}]","workFenceName":"经营-商圈-印保利广场右下角-B1南市街-晚高峰"}}}
