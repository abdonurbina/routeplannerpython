import pandas as pd
from ortools.constraint_solver import pywrapcp, routing_enums_pb2


def read_excel(file_path):
    """Read stop data from an Excel file."""
    data = pd.read_excel(file_path)
    print(data.head())
    return data[['Stop ID', 'Latitude', 'Longitude']].values


def calculate_distance_matrix(locations):
    """Calculate the distance matrix using haversine formula."""
    from math import radians, cos, sin, sqrt, atan2

    def haversine(lat1, lon1, lat2, lon2):
        R = 6371  # Earth radius in kilometers
        dlat = radians(lat2 - lat1)
        dlon = radians(lon2 - lon1)
        a = sin(dlat / 2) ** 2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon / 2) ** 2
        c = 2 * atan2(sqrt(a), sqrt(1 - a))
        return R * c

    size = len(locations)
    distance_matrix = [[0] * size for _ in range(size)]
    for i in range(size):
        for j in range(size):
            if i != j:
                distance_matrix[i][j] = haversine(
                    locations[i][1], locations[i][2],
                    locations[j][1], locations[j][2]
                )
    return distance_matrix


def create_route_plan(distance_matrix, num_vehicles, depot=0):
    """Solve the routing problem and create a route plan."""
    manager = pywrapcp.RoutingIndexManager(len(distance_matrix), num_vehicles, depot)
    routing = pywrapcp.RoutingModel(manager)

    def distance_callback(from_index, to_index):
        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        return distance_matrix[from_node][to_node]

    transit_callback_index = routing.RegisterTransitCallback(distance_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)

    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = (
        routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC)

    solution = routing.SolveWithParameters(search_parameters)
    if solution:
        return extract_routes(solution, routing, manager)
    else:
        return None


def extract_routes(solution, routing, manager):
    """Extract the route plan from the solution."""
    routes = []
    for vehicle_id in range(routing.vehicles()):
        index = routing.Start(vehicle_id)
        route = []
        while not routing.IsEnd(index):
            route.append(manager.IndexToNode(index))
            index = solution.Value(routing.NextVar(index))
        route.append(manager.IndexToNode(index))  # Add the depot
        routes.append(route)
    return routes


# Main script
if __name__ == "__main__":
    # Replace 'stops.xlsx' with your actual Excel file
    file_path = 'routes.xlsx'
    num_vehicles = 1  # Adjust based on your fleet size
    stops = read_excel(file_path)

    distance_matrix = calculate_distance_matrix(stops)
    routes = create_route_plan(distance_matrix, num_vehicles)

    if routes:
        print("Route plan:")
        for i, route in enumerate(routes):
            print(f"Vehicle {i + 1}: {route}")
    else:
        print("No solution found.")
