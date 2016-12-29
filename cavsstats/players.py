from cavsstats.soup_utils import getSoupFromURL
import re
import logging
import json

class Player(object):
    first_names = []

    def __init__(self, team):