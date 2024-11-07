def checkMannWitney(transport_ecology, non_transport_ecology):
    from scipy.stats import mannwhitneyu
    # Perform the Mann-Whitney U Test
    stat, p_value = mannwhitneyu(transport_ecology, non_transport_ecology, alternative='less')

    print(f"Mann-Whitney U statistic: {stat}")
    print(f"P-value: {p_value}")

transport_n1 = "2	3	2	5	4	5	4	3	5	4	3	2	5	4	3	4	5	3	1	5	3	3	5	4	3	1	3	4	1	3	2	4	3	5	3	1	2	5	3	5	3	2	4	3	2	3	2	3	4	4	4	5	3	4	2	1	2	2	4	1	5	2	1	2".split()
transport_n2 = "2	3	1	3	5	4	2	3	5	4	4	3	4	5	3	4	5	3	1	4	2	3	5	4	3	1	3	3	1	3	2	3	3	5	4	1	4	5	3	5	5	2	4	4	3	3	3	3	4	4	5	5	3	4	3	3	4	2	5	1	4	2	1	3".split()
transport_n3 = "3	3	1	3	5	4	3	3	5	4	3	3	4	4	3	5	5	3	1	4	3	3	5	4	3	1	2	2	1	3	2	4	3	5	4	1	2	5	5	5	5	2	4	4	2	4	2	3	5	4	4	5	3	4	3	2	4	2	5	1	3	2	1	2".split()
transport_n4 = "2	3	2	3	5	5	2	3	5	4	3	2	4	3	3	5	5	3	1	4	4	4	5	4	3	1	1	3	1	3	2	4	3	4	4	1	4	5	5	5	5	2	4	4	3	3	3	2	4	4	5	4	3	4	3	2	3	1	4	1	4	2	1	2".split()

transport_d1 = "5	3	4	2	2	4	1	1	1	2	5	4	2	3	2	3	4	1	1	2	1	2	3	4	1	1	1	3	2	1	4	1	1	1	1	2	3	1	2	2	2	4	2	2	2	1	4	2	3	2	2	3	1	3	1	2	3	1	1	2	2	4	2	3	2	4	4	3	1	3	1	2	1	2	2	3	3	1	1	4	3	1	1	4".split()
transport_d2 = "5	3	3	1	2	4	1	1	1	2	5	5	2	2	1	2	4	2	1	4	1	2	3	4	1	2	4	4	2	2	3	3	1	1	1	3	2	2	2	2	2	4	2	1	1	1	3	1	4	3	2	3	1	1	1	3	3	1	1	2	3	4	2	4	2	4	4	3	1	3	1	1	2	3	2	3	2	1	1	3	3	1	1	3".split()
transport_d3 = "5	3	3	2	2	4	1	1	1	2	5	5	2	2	1	2	4	2	1	4	2	1	3	5	1	2	4	4	2	2	3	3	1	1	1	3	2	2	2	2	2	3	2	1	1	1	3	2	3	1	3	3	1	1	1	3	3	1	1	2	3	4	2	4	2	4	4	3	2	3	1	1	2	3	2	3	2	1	1	3	3	1	1	3".split()
transport_d4 = "2	3	3	1	2	4	1	1	1	2	5	5	2	3	1	1	4	1	1	3	2	1	3	3	1	2	3	4	2	2	2	1	1	1	1	2	3	1	2	2	2	3	2	2	1	1	4	1	4	2	3	3	1	2	1	3	3	1	1	2	3	4	3	3	2	2	5	2	1	3	1	1	2	3	2	3	3	1	1	3	3	3	1	4".split()

transport_n1 = [*map(int,transport_n1)]
transport_d1 = [*map(int,transport_d1)]
transport_n2 = [*map(int,transport_n2)]
transport_d2 = [*map(int,transport_d2)]
transport_n3 = [*map(int,transport_n3)]
transport_d3 = [*map(int,transport_d3)]
transport_n4 = [*map(int,transport_n4)]
transport_d4 = [*map(int,transport_d4)]
print(transport_n1)
checkMannWitney(transport_d1, transport_n1)
checkMannWitney(transport_d2, transport_n2)
checkMannWitney(transport_d3, transport_n3)
checkMannWitney(transport_d4, transport_n4)

