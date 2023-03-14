from testcase import TestCase

class WasRun(TestCase):
    def __init__(self, name):
        self.wasRun = None
        super().__init__(name)
    def run(self):
        method = getattr(self, self.name)
        method()
    def testMethod(self):
        self.wasRun = 1