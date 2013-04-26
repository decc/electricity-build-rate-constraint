data = {}
request = undefined
s =
  version: 1
  build_rate: 0

url_structure = [
  "version"
  "build_rate"
]

update = () ->
  if request?
    setUrlToMatchSettings()
    request.abort()
  request = d3.json("/data#{urlForSettings()}", (error, json) ->
    return console.warn(error) if error
    data = json
    window.data = data
    visualise()
  )

setUrlToMatchSettings = () ->
  history.pushState(undefined, undefined, urlForSettings())

window.onpopstate = () ->
  getSettingsFromUrl()
  updateControlsFromSettings()
  update()

urlForSettings = () ->
  url = for a in url_structure
    s[a]
  "/" + url.join('/')

getSettingsFromUrl = () ->
  c = window.location.pathname.split('/')
  c.shift()

  for a,i in url_structure
    if c[i]? and c[i] != ""
      s[a] = c[i]

  

timeSeriesChart = ->
  
  margin =
    top: 20
    right: 20
    bottom: 20
    left: 40

  width = 240
  height = 120

  X = (d) -> xScale(d[0])
  Y = (d) -> yScale(d[1])
  xValue = (d) -> d[0]
  yValue = (d) -> d[1]
  xScale = d3.time.scale()
  yScale = d3.scale.linear()
  xAxis = d3.svg.axis().scale(xScale).orient("bottom")
  yAxis = d3.svg.axis().scale(yScale).orient("left")
  area = d3.svg.area().x(X).y1(Y)
  line = d3.svg.line().x(X).y(Y)

  chart = (selection) ->
    selection.each (data) ->
      
      # Convert data to standard representation greedily;
      # this is needed for nondeterministic accessors.
      data = data.map((d, i) ->
        [xValue.call(data, d, i), yValue.call(data, d, i)]
      )
      # Update the x-scale.
      xScale.domain(d3.extent(data, (d) -> d[0]))
        .range([0, width - margin.left - margin.right])
      
      # Update the y-scale.
      yScale.domain([0, d3.max(data, (d) -> d[1])])
        .range([height - margin.top - margin.bottom, 0])
      
      # Select the svg element, if it exists.
      svg = d3.select(this).selectAll("svg").data([data])
      
      # Otherwise, create the skeletal chart.
      gEnter = svg.enter()
        .append("svg")
        .append("g")

      # And the basic bits
      gEnter
        .append("path")
        .attr("class", "area")
      gEnter
        .append("path")
        .attr("class", "line")
      gEnter
        .append("g")
        .attr("class", "x axis")
      gEnter
        .append("g")
        .attr("class", "y axis")
      
      # Update the outer dimensions.
      svg
        .attr("width", width)
        .attr("height", height)
      
      # Update the inner dimensions.
      g = svg.select("g").attr("transform", "translate(" + margin.left + "," + margin.top + ")")
      
      # Update the area path.
      g.select(".area").attr("d", area.y0(yScale.range()[0]))
      
      # Update the line path.
      g.select(".line").attr("d", line)
      
      # Update the x-axis.
      g.select(".x.axis")
        .attr("transform", "translate(0," + yScale.range()[0] + ")")
        .call(xAxis)
       
      # Update the y-axis.
      g.select(".y.axis")
        .attr("transform", "translate(0," + xScale.range()[0] + ")")
        .call(yAxis)

  chart.margin = (_) ->
    return margin  unless _?
    margin = _
    chart

  chart.width = (_) ->
    return width  unless _?
    width = _
    chart

  chart.height = (_) ->
    return height  unless _?
    height = _
    chart

  chart.x = (_) ->
    return xValue  unless _?
    xValue = _
    chart

  chart.y = (_) ->
    return yValue  unless _?
    yValue = _
    chart

  chart


chart = timeSeriesChart()
  .x((d,i) -> new Date("#{2012+i}"))
  .y((d) -> +d)

visualise = () ->
  d3.select('#zero_carbon_build_rate')
    .datum(data.series.zero_carbon_build_rate)
    .call(chart)

  d3.select('#emissions')
    .datum(data.series.emissions)
    .call(chart)

  d3.select('#emissions_factor')
    .datum(data.series.emissions_factor)
    .call(chart)

d3.select('#maximum_low_carbon_build_rate')
  .on('change', () ->
    d3.select('#maximum_low_carbon_build_rate_value').text(+this.value)
    s.build_rate = +this.value
    update()
  )

updateControlsFromSettings = () ->
  d3.select('#maximum_low_carbon_build_rate').property('value',s.build_rate)
  d3.select('#maximum_low_carbon_build_rate_value').text(s.build_rate)

getSettingsFromUrl()
updateControlsFromSettings()
update()
