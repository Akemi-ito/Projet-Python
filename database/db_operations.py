import pymongo
from pymongo import MongoClient
import json

class DBOperations:
    def __init__(self, db_name="spacex_db", collection_name="launches"):
        self.uri = "mongodb://localhost:27017/" #lien connection mongo db
        self.client = MongoClient(self.uri)
        self.db = self.client[db_name]
        self.collection = self.db[collection_name]

    def insert_data(self, data):
        """Insert les données"""
        if isinstance(data, list):
            self.collection.insert_many(data)
        else:
            self.collection.insert_one(data)
        return True

    def clear_collection(self):
        """drop les données"""
        self.collection.delete_many({})
        return True

    def get_launches_count(self):
        """recupere le nombre de lancements dans la collection"""
        return self.collection.count_documents({})

    def get_launches_by_type(self):
        """recupere les lancements"""
        pipeline = [
            {"$group": {"_id": "$mission_name", "count": {"$sum": 1}}},
            {"$sort": {"count": -1}}
        ]
        return list(self.collection.aggregate(pipeline))
    def get_launches_by_organization(self):
        pipeline = [
        {
            "$project": {
                "mission_name": 1,
                # Catégorisation des missions
                "organization": {
                    "$switch": {
                        "branches": [
                            # Missions Starlink
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "STARLINK|STARLINK MISSION: SPACEX’S 100TH SUCCESSFUL FLIGHT", "options": "i"}},
                                "then": "Starlink"
                            },
                            # Missions Crew Dragon (vols habités)
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "^CREW[- ]([1-9]|10)( MISSION)?$", "options": "i"}},
                                "then": "NASA/Vols habités (Crew Dragon)"
                            },
                            # Missions cargo Dragon
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "CRS", "options": "i"}},
                                "then": "NASA/Cargo"
                            },
                            # Missions Starship
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "STARSHIP FLIGHT TEST|^STARSHIP'S", "options": "i"}},
                                "then": "Starship Tests"
                            },
                            # Missions NROL (reconnaissance)
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "NROL", "options": "i"}},
                                "then": "NRO (Reconnaissance)"
                            },
                            # Missions GPS
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "GPS", "options": "i"}},
                                "then": "USAF/Space Force (GPS)"
                            },
                            # Missions USSF
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "USSF", "options": "i"}},
                                "then": "US Space Force"
                            },
                            # Missions OneWeb
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "ONEWEB", "options": "i"}},
                                "then": "OneWeb"
                            },
                            # Missions SES
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "SES|O3B", "options": "i"}},
                                "then": "SES"
                            },
                            # Missions Intelsat
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "INTELSAT", "options": "i"}},
                                "then": "Intelsat"
                            },
                            # Missions Eutelsat
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "EUTELSAT", "options": "i"}},
                                "then": "Eutelsat"
                            },
                            # Missions Transporter (rideshare)
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "TRANSPORTER", "options": "i"}},
                                "then": "Rideshare (Transporter)"
                            },
                            # Missions Maxar
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "MAXAR", "options": "i"}},
                                "then": "Maxar"
                            },
                            # Missions Bandwagon
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "BANDWAGON", "options": "i"}},
                                "then": "Bandwagon"
                            },
                            # Missions ESA
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "ESA|EUCLID|EARTHCARE", "options": "i"}},
                                "then": "ESA"
                            },
                            # Autres missions gouvernementales
                            {
                                "case": {"$regexMatch": {"input": "$mission_name", "regex": "NASA|DSCOVR|SWOT|PACE|IXPE|DART|KPLO|SENTINEL|SAOCOM", "options": "i"}},
                                "then": "Autres agences spatiales"
                            }
                        ],
                        "default": "Autres clients commerciaux"
                    }
                }
            }
        },
        {
            "$group": {
                "_id": "$organization",
                "count": {"$sum": 1},
                "missions": {"$push": "$mission_name"}
            }
        },
        {
            "$sort": {"count": -1}
        }
    ]
        return list(self.collection.aggregate(pipeline))