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
    left: 50

  width = 250
  height = 125

  unit = "TWh/yr"
  first_scale_year = 2010
  last_scale_year = 2050
  first_data_year = 2012
  max_value = undefined

  xScale = d3.time.scale()
  yScale = d3.scale.linear()

  xAxis = d3.svg.axis().scale(xScale).orient("bottom").ticks(5)
  yAxis = d3.svg.axis().scale(yScale).orient("left").ticks(5)

  area = d3.svg.area()
    .x((d,i) -> xScale(new Date("#{first_data_year+i}")))
    .y((d,i) -> yScale(+d))

  line = d3.svg.line()
    .x((d,i) -> xScale(new Date("#{first_data_year+i}")))
    .y((d,i) -> yScale(+d))

  chart = (selection) ->
    selection.each (data) ->
      
      # Update the x-scale.
      xScale
        .domain([new Date("#{first_scale_year}"), new Date("#{last_scale_year}")])
        .range([0, width - margin.left - margin.right])
      
      # Update the y-scale.
      yScale
        .domain([0, max_value || d3.max(data)])
        .range([height - margin.top - margin.bottom, 0])
      
      # Select the svg element, if it exists.
      svg = d3.select(this).selectAll("svg").data([data])
      
      # Otherwise, create the skeletal chart.
      gEnter = svg.enter()
        .append("svg")
        .append("g")
        .attr('class','drawing')

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
      gEnter
        .append("text")
        .attr("class", "y axislabel")
      
      # Update the outer dimensions.
      svg
        .attr("width", width)
        .attr("height", height)
      
      # Update the inner dimensions.
      g = svg.select("g").attr("transform", "translate(" + margin.left + "," + margin.top + ")")
      
      # Update the area path.
      g.select(".area")
        .transition()
          .attr("d", area.y0(yScale.range()[0]))
      
      # Update the line path.
      g.select(".line")
        .transition()
          .attr("d", line)
      
      # Update the x-axis.
      g.select(".x.axis")
        .attr("transform", "translate(0," + yScale.range()[0] + ")")
        .call(xAxis)
       
      # Update the y-axis.
      g.select(".y.axis")
        .attr("transform", "translate(0," + xScale.range()[0] + ")")
        .call(yAxis)

      g.select(".y.axislabel")
        .attr("transform", "translate(0," + (xScale.range()[0] - 10) + ")")
        .text(unit)

  chart.margin = (_) ->
    return margin unless _?
    margin = _
    chart

  chart.width = (_) ->
    return width unless _?
    width = _
    chart

  chart.height = (_) ->
    return height unless _?
    height = _
    chart

  chart.unit = (_) ->
    return unit unless _?
    unit = _
    chart

  chart.first_scale_year = (_) ->
    return first_scale_year unless _?
    first_scale_year = _
    chart

  chart.last_scale_year = (_) ->
    return last_scale_year unless _?
    last_scale_year = _
    chart

  chart.first_data_year = (_) ->
    return first_data_year unless _?
    first_data_year = _
    chart

  chart.max_value = (_) ->
    return max_value unless _?
    max_value = _
    chart

  chart


visualise = () ->
  d3.select('#zero_carbon_build_rate')
    .datum(data.series.zero_carbon_build_rate)
    .call(timeSeriesChart().unit("TWh/yr/yr").max_value(100))

  d3.select('#emissions')
    .datum(data.series.emissions)
    .call(timeSeriesChart().unit("MtCO2/yr").max_value(200))

  d3.select('#emissions_factor')
    .datum(data.series.emissions_factor)
    .call(timeSeriesChart().unit("gCO2/kWh").max_value(500))

  d3.select('#emissions_factor .chart svg')

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
