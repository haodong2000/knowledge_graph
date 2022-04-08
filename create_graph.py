# -*- utf-8 -*-

from neo4j import GraphDatabase
import logging
from neo4j.exceptions import ServiceUnavailable
import data_preprocess


class App:

    def __init__(self, uri, user, password):
        self.driver = GraphDatabase.driver(uri, auth=(user, password))
        self.all_subjects = []
        self.all_objects = []

    def check_current(self, current_sub, current_ob):
        flag_1, flag_2 = True, True
        for name in self.all_subjects:
            if name == current_sub:
                flag_1 = False
        for name in self.all_objects:
            if name == current_ob:
                flag_2 = False
        self.all_subjects.append(current_sub)
        self.all_objects.append(current_ob)
        return flag_1, flag_2

    def close(self):
        # Don't forget to close the driver connection when you are finished with it
        self.driver.close()

    def create_relation(self, piece):
        subject, sub_type, relation, object_, obj_type = piece[
            0], piece[1], piece[2], piece[3], piece[4]
        flag_1, flag_2 = self.check_current(subject, object_)
        with self.driver.session() as session:
            # Write transactions allow the driver to handle retries and transient errors
            result = session.write_transaction(
                self._create_and_return_nodes, subject, sub_type, relation, object_, obj_type, flag_1, flag_2)
            for row in result:
                print("Created relation ->", row, flag_1, flag_2)
        with self.driver.session() as session:
            # Write transactions allow the driver to handle retries and transient errors
            _ = session.write_transaction(
                self._create_relation, subject, sub_type, relation, object_, obj_type, flag_1, flag_2)

    @staticmethod
    def _create_relation(tx, subject, sub_type, relation, object_, obj_type, flag_1, flag_2):
        query = (
            "MATCH (a: " + sub_type + "), (b: " + obj_type + ") WHERE a.name=$subject AND b.name=$object_ CREATE (a)-[:" + relation + "] -> (b)"
        )
        _ = tx.run(query, sub_type=sub_type, obj_type=obj_type, subject=subject, object_=object_)

    @staticmethod
    def _create_and_return_nodes(tx, subject, sub_type, relation, object_, obj_type, flag_1, flag_2):
        # To learn more about the Cypher syntax, see https://neo4j.com/docs/cypher-manual/current/
        # The Reference Card is also a good resource for keywords https://neo4j.com/docs/cypher-refcard/current/
        if flag_1 and flag_2:
            query = (
                "CREATE (" + subject + ":" + sub_type + " { name: $subject }) "
                "CREATE (" + object_ + ":" + obj_type + " { name: $object_ }) "
                "RETURN " + subject + ", " + object_
            )
        elif flag_1 is False and flag_2 is True:
            query = (
                "CREATE (" + object_ + ":" + obj_type + " { name: $object_ }) "
                "RETURN " + object_
            )
        elif flag_2 is False and flag_1 is True:
            query = (
                "CREATE (" + subject + ":" + sub_type + " { name: $subject }) "
                "RETURN " + subject
            )
        else:
            query = (
                "RETURN 0"
            )
        result = tx.run(query, subject=subject, sub_type=sub_type, object_=object_, obj_type=obj_type)
        try:
            if flag_1 and flag_2:
                return [{subject: row[subject]["name"], object_: row[object_]["name"]}
                        for row in result]
            elif flag_1 is False and flag_2 is True:
                return [{object_: row[object_]["name"]}
                        for row in result]
            elif flag_2 is False and flag_1 is True:
                return [{subject: row[subject]["name"]}
                        for row in result]
            else:
                return ["only relation" for row in result]
        # Capture any errors along with the query and data for traceability
        except ServiceUnavailable as exception:
            logging.error("{query} raised an error: \n {exception}".format(
                query=query, exception=exception))
            raise

    def find_object(self, object_, obj_type):
        with self.driver.session() as session:
            result = session.read_transaction(
                self._find_and_return_object, object_, obj_type)
            for row in result:
                print("Found object ->", row)

    @staticmethod
    def _find_and_return_object(tx, object_, obj_type):
        query = (
            "MATCH (p:" + obj_type + ") "
            "WHERE p.name = $object_ "
            "RETURN p.name AS name"
        )
        result = tx.run(query, object_=object_, obj_type=obj_type)
        return [row["name"] for row in result]

    def delete_all_graph(self):
        with self.driver.session() as session:
            result = session.write_transaction(self._delete_and_return_graph)
            for row in result:
                print("Delete result ->", row)

    @staticmethod
    def _delete_and_return_graph(tx):
        query = (
            "MATCH (n)"
            "DETACH DELETE n"
        )
        result = tx.run(query)
        return [row for row in result]


if __name__ == "__main__":
    # Aura queries use an encrypted connection using the "neo4j+s" URI scheme
    uri = "neo4j+s://e8710434.databases.neo4j.io"
    user = "neo4j"
    password = ""
    filename = "./data/en_relation.xls"
    data = data_preprocess.read_data(filename, sheet_name="en_relation")
    app = App(uri, user, password)
    app.delete_all_graph()
    for i in range(len(data)):
        piece = [data[data.columns[j]][i] for j in range(5)]
        app.create_relation(piece)
    app.find_object("AKI", "DISEASE")
    app.close()
