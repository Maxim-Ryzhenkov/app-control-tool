# coding: utf-8

from packaging import version


class ApplicationVersion:
    def __init__(self, version_string):
        self.string = version_string
        self.version = version.parse(self.string)
        self.major = None
        self.minor = None

    def __eq__(self, other):
        return self.version == version.parse(other)

    def __ne__(self, other):
        return self.version != version.parse(other)

    def __lt__(self, other):
        return self.version < version.parse(other)

    def __le__(self, other):
        return self.version <= version.parse(other)

    def __gt__(self, other):
        return self.version > version.parse(other)

    def __ge__(self, other):
        return self.version >= version.parse(other)

    def __repr__(self):
        return f"Version ({self.string})"
